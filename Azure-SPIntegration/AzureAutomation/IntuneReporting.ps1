<#
.SYNOPSIS
    Exports Entra devices with RegisteredOwners and compliance status (Azure Automation).

.DESCRIPTION
    Uses Get-MgDevice -All, resolves RegisteredOwners via Graph batch,
    exports Devices + DeviceSummary CSVs, then uploads them to SharePoint Documents
    via Connect-PnPOnline -ManagedIdentity.

.NOTES
    Azure Automation setup
    ----------------------
    1. Enable system-assigned Managed Identity on the Automation Account.
    2. Grant Graph application permissions (admin consent):
         Device.Read.All
         Directory.Read.All
         User.Read.All
         Sites.ReadWrite.All   (or SharePoint Sites.FullControl.All for PnP)
       Re-run Set-SystemManagedId.ps1 after updating permissions, then wait a few minutes.
    3. Runtime modules (same Graph version, e.g. 2.38.1):
         Microsoft.Graph.Authentication
         Microsoft.Graph.Users
         Microsoft.Graph.Identity.DirectoryManagement
         PnP.PowerShell
    4. Automation Variables:
         SHAREPOINT_SITE_URL
         SHAREPOINT_FOLDER_PATH_INTUNE  (optional; default Shared Documents/IntuneReports)

.EXAMPLE
    .\IntuneReporting.ps1
#>
[CmdletBinding()]
Param(
    [string]$ExportPath = (Join-Path $env:TEMP 'Intune-Report')
)

$ErrorActionPreference = 'Stop'
$BatchSize = 10
$BatchDelayMs = 1000
$GraphMaxRetries = 8
$GraphBaseDelaySeconds = 10
$runStart = Get-Date

$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
$SharePointFolderPath = Get-AutomationVariable -Name 'SHAREPOINT_FOLDER_PATH_INTUNE' -ErrorAction SilentlyContinue

if ([string]::IsNullOrWhiteSpace($SharePointFolderPath)) {
    $SharePointFolderPath = 'Shared Documents/IntuneReports'
}

if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
    throw 'Automation variable SHAREPOINT_SITE_URL is required (e.g. https://contoso.sharepoint.com/sites/IT).'
}

function Write-RunbookLog {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Output "[$stamp] $Message"
}

function Test-GraphThrottleError {
    param($ErrorRecord)

    if (-not $ErrorRecord) { return $false }

    $message = [string]$ErrorRecord.Exception.Message
    if ($message -match '429|TooManyRequests|Too Many Requests') {
        return $true
    }

    $status = $null
    if ($ErrorRecord.Exception.Response) {
        try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch { }
    }
    return $status -eq 429
}

function Get-GraphRetryDelaySeconds {
    param(
        $ErrorRecord,
        [int]$Attempt,
        [int]$BaseDelaySeconds = 10
    )

    $message = [string]$ErrorRecord.Exception.Message
    if ($message -match 'Retry-After\s*:\s*(\d+)') {
        return [Math]::Max(1, [int]$Matches[1])
    }

    if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Headers) {
        try {
            $header = $ErrorRecord.Exception.Response.Headers['Retry-After']
            if ($header) {
                return [Math]::Max(1, [int]("$header"))
            }
        }
        catch { }
    }

    return [Math]::Min(120, $BaseDelaySeconds * $Attempt)
}

function Invoke-WithGraphRetry {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [string]$Activity = 'Graph request',

        [int]$MaxRetries = 8,

        [int]$BaseDelaySeconds = 10
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ((Test-GraphThrottleError -ErrorRecord $_) -and $attempt -lt $MaxRetries) {
                $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt -BaseDelaySeconds $BaseDelaySeconds
                Write-RunbookLog "Throttled during '$Activity' (attempt $attempt/$MaxRetries). Waiting $delay second(s)..."
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Import-MatchingGraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    $authVersions = @(
        Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
            Select-Object -ExpandProperty Version -Unique |
            Sort-Object -Descending
    )

    if ($authVersions.Count -eq 0) {
        throw 'Microsoft.Graph.Authentication is not installed in this Automation Account.'
    }

    $commonVersion = $null
    foreach ($version in $authVersions) {
        $missing = @()
        foreach ($moduleName in $requiredModules) {
            $match = Get-Module -ListAvailable -Name $moduleName |
                Where-Object { $_.Version -eq $version } |
                Select-Object -First 1
            if (-not $match) {
                $missing += $moduleName
            }
        }

        if ($missing.Count -eq 0) {
            $commonVersion = $version
            break
        }
    }

    if (-not $commonVersion) {
        $installed = foreach ($moduleName in $requiredModules) {
            $versions = @(
                Get-Module -ListAvailable -Name $moduleName |
                    Select-Object -ExpandProperty Version -Unique |
                    Sort-Object -Descending
            )
            "  $moduleName : $(if ($versions) { $versions -join ', ' } else { '(not installed)' })"
        }

        throw @"
Microsoft Graph modules must all be the SAME version in Azure Automation.
Install matching versions of:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Users
  - Microsoft.Graph.Identity.DirectoryManagement

Currently available:
$($installed -join [Environment]::NewLine)
"@
    }

    foreach ($moduleName in $requiredModules) {
        Import-Module $moduleName -RequiredVersion $commonVersion -Force -ErrorAction Stop
    }

    Write-RunbookLog "Loaded Microsoft Graph modules version $commonVersion"
}

function Connect-MgGraphManagedIdentity {
    Import-MatchingGraphModules

    Write-RunbookLog 'Connecting to Microsoft Graph with system-assigned Managed Identity...'
    Connect-MgGraph -Identity -NoWelcome

    if (Get-Command Set-MgRequestContext -ErrorAction SilentlyContinue) {
        Set-MgRequestContext -Retries $GraphMaxRetries -RetryDelay $GraphBaseDelaySeconds | Out-Null
    }

    $context = Get-MgContext
    Write-RunbookLog "Connected. AuthType=$($context.AuthType); TenantId=$($context.TenantId); AppName=$($context.AppName)"
    if ($context.Scopes) {
        Write-RunbookLog "Token roles/scopes: $($context.Scopes -join ', ')"
        $needed = @('Device.Read.All', 'Directory.Read.All')
        $missing = @($needed | Where-Object { $context.Scopes -notcontains $_ })
        if ($missing.Count -eq $needed.Count) {
            Write-RunbookLog "WARNING: Token does not include Device.Read.All or Directory.Read.All. Permissions may be on a different enterprise app, or the token is stale. Confirm Enterprise application = '$($context.AppName)' and wait 5-10 minutes after assigning roles."
        }
    }
    else {
        Write-RunbookLog 'WARNING: Get-MgContext.Scopes is empty; cannot verify Device.Read.All in the access token.'
    }
}

function Get-GraphCollectionValue {
    param($Page)

    if ($null -eq $Page) { return @() }
    if ($Page -is [System.Collections.IDictionary]) {
        return @($Page['value'])
    }
    return @($Page.value)
}

function Get-GraphNextLink {
    param($Page)

    if ($null -eq $Page) { return $null }
    if ($Page -is [System.Collections.IDictionary]) {
        return $Page['@odata.nextLink']
    }
    return $Page.'@odata.nextLink'
}

function Get-GraphProp {
    param($Object, [string]$Name)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]::Equals([string]$key, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                return $Object[$key]
            }
        }
        return $null
    }

    foreach ($prop in $Object.PSObject.Properties) {
        if ([string]::Equals($prop.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) {
            return $prop.Value
        }
    }
    return $null
}

function Invoke-MgGraphBatch {
    param(
        [Parameter(Mandatory)]
        [array]$Requests,
        [int]$MaxRetries = 6
    )

    $body = @{ requests = $Requests } | ConvertTo-Json -Depth 6

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $result = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body
            return @($result.responses)
        }
        catch {
            if ((Test-GraphThrottleError -ErrorRecord $_) -and $attempt -lt $MaxRetries) {
                $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt -BaseDelaySeconds $GraphBaseDelaySeconds
                Write-RunbookLog "Batch throttled (attempt $attempt/$MaxRetries). Waiting $delay second(s)..."
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Invoke-MgGraphBatchPages {
    param(
        [Parameter(Mandatory)]
        [array]$AllRequests,
        [int]$Size = 10,
        [int]$DelayMs = 1000,
        [string]$Activity = 'Graph batch'
    )

    $allResponses = [System.Collections.Generic.List[object]]::new()
    if ($AllRequests.Count -eq 0) { return @() }

    $totalBatches = [Math]::Ceiling($AllRequests.Count / [double]$Size)

    for ($i = 0; $i -lt $AllRequests.Count; $i += $Size) {
        $batchNum = [int]($i / $Size) + 1
        $end = [Math]::Min($i + $Size - 1, $AllRequests.Count - 1)
        $chunk = @($AllRequests[$i..$end])

        Write-RunbookLog "$Activity - batch $batchNum of $totalBatches"
        $batchResponses = Invoke-MgGraphBatch -Requests $chunk -MaxRetries $GraphMaxRetries

        $finalResponses = [System.Collections.Generic.List[object]]::new()
        foreach ($response in @($batchResponses)) {
            if ([int]$response.status -eq 429) {
                $retryId = [string]$response.id
                $original = $chunk | Where-Object { [string]$_.id -eq $retryId } | Select-Object -First 1
                if ($original) {
                    Write-RunbookLog "$Activity - retrying throttled item id=$retryId"
                    Start-Sleep -Seconds $GraphBaseDelaySeconds
                    $retryResult = Invoke-MgGraphBatch -Requests @($original) -MaxRetries $GraphMaxRetries
                    $finalResponses.AddRange(@($retryResult))
                    continue
                }
            }
            $finalResponses.Add($response)
        }

        $allResponses.AddRange($finalResponses)

        if ($DelayMs -gt 0 -and ($i + $Size) -lt $AllRequests.Count) {
            Start-Sleep -Milliseconds $DelayMs
        }
    }

    return @($allResponses)
}

function Get-DeviceRegisteredOwnerIds {
    param(
        [array]$Devices,
        [int]$Size,
        [int]$DelayMs
    )

    $map = @{}
    foreach ($d in $Devices) {
        $map[$d.Id] = [System.Collections.Generic.List[string]]::new()
    }

    $requests = for ($i = 0; $i -lt $Devices.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/devices/$($Devices[$i].Id)/registeredOwners?`$select=id" }
    }

    Write-RunbookLog "Loading RegisteredOwners for $($Devices.Count) device(s)..."
    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'RegisteredOwners'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $deviceId = $Devices[$index].Id
        if ($response.status -eq 200) {
            foreach ($obj in @($response.body.value)) {
                if ($obj.id) { $map[$deviceId].Add([string]$obj.id) }
            }
        }
    }

    return $map
}

function Import-UsersById {
    param(
        [string[]]$UserIds,
        [int]$Size,
        [int]$DelayMs
    )

    $cache = @{}
    $unique = @($UserIds | Where-Object { $_ } | Select-Object -Unique)
    if ($unique.Count -eq 0) { return $cache }

    Write-RunbookLog "Resolving $($unique.Count) RegisteredOwner(s)..."
    $requests = for ($i = 0; $i -lt $unique.Count; $i++) {
        @{
            id     = "$i"
            method = 'GET'
            url    = "/users/$($unique[$i])?`$select=id,displayName,userPrincipalName"
        }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Resolve owners'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $userId = $unique[$index]
        if ($response.status -eq 200) {
            $cache[$userId] = [PSCustomObject]@{
                DisplayName       = $response.body.displayName
                UserPrincipalName = $response.body.userPrincipalName
            }
        }
    }

    return $cache
}

function Get-PnPDocumentFolderPath {
    param([string]$FolderPath)

    $normalized = ($FolderPath -replace '\\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return 'Shared Documents'
    }

    if ($normalized -match '^(Shared Documents|Documents)(/|$)') {
        return $normalized
    }

    return "Shared Documents/$normalized"
}

function Get-StableReportFileName {
    param([Parameter(Mandatory)][string]$FilePath)

    $leaf = Split-Path -Path $FilePath -Leaf
    switch -Regex ($leaf) {
        'DeviceSummary' { return 'Intune-Report-DeviceSummary.csv' }
        'Devices-' { return 'Intune-Report-Devices.csv' }
        default { return $leaf }
    }
}

function Publish-ReportToSharePoint {
    param(
        [Parameter(Mandatory)]
        [string[]]$FilePaths,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [string]$FolderPath = 'Shared Documents/IntuneReports'
    )

    Import-Module PnP.PowerShell -ErrorAction Stop

    $folder = Get-PnPDocumentFolderPath -FolderPath $FolderPath
    Write-RunbookLog "Uploading reports to SharePoint site '$SiteUrl' (folder: '$folder') via Connect-PnPOnline..."

    Connect-PnPOnline -Url $SiteUrl -ManagedIdentity

    try {
        Resolve-PnPFolder -SiteRelativePath $folder | Out-Null

        foreach ($filePath in $FilePaths) {
            if (-not (Test-Path -LiteralPath $filePath)) {
                throw "Report file not found: $filePath"
            }

            $fileName = Get-StableReportFileName -FilePath $filePath
            Add-PnPFile -Path $filePath -Folder $folder -NewFileName $fileName | Out-Null
            Write-RunbookLog "Uploaded/updated SharePoint file: $folder/$fileName"
        }
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

try {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $exportRoot = Join-Path $env:TEMP ("Intune-Report-" + $timestamp)
    if (-not (Test-Path -LiteralPath $exportRoot)) {
        New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
    }

    $devicesPath = Join-Path $exportRoot "Devices-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "DeviceSummary-$timestamp.csv"

    Write-RunbookLog 'Starting Intune/Entra device report'
    Connect-MgGraphManagedIdentity

    Write-RunbookLog 'Retrieving devices (paged) with registeredOwners expand...'
    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $compliantCount = 0
    $nonCompliantCount = 0
    $unknownComplianceCount = 0
    $noOwnerCount = 0
    $resolvedOwnerIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    $pageUri = 'https://graph.microsoft.com/v1.0/devices?$select=id,deviceId,displayName,isCompliant,accountEnabled,operatingSystem&$expand=registeredOwners($select=id,displayName,userPrincipalName)&$top=100'
    $pageNumber = 0

    while (-not [string]::IsNullOrWhiteSpace($pageUri)) {
        $pageNumber++
        try {
            $page = Invoke-WithGraphRetry -Activity "Get devices page $pageNumber" -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
                Invoke-MgGraphRequest -Method GET -Uri $pageUri
            }
        }
        catch {
            if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges') {
                throw @"
Graph denied access to /devices (page $pageNumber).
AppName from token: $((Get-MgContext).AppName)
Roles/scopes: $((Get-MgContext).Scopes -join ', ')

Confirm in Entra > Enterprise applications > aa-automation > Permissions:
  - Device.Read.All (Application)
  - Directory.Read.All (Application)
Then wait 5-10 minutes and retry. Also confirm Set-SystemManagedId.ps1 PrincipalId is aa-automation's Object ID.
Original error: $($_.Exception.Message)
"@
            }
            throw
        }

        $pageDevices = Get-GraphCollectionValue -Page $page
        foreach ($d in $pageDevices) {
            if ($null -eq $d) { continue }

            $ownersRaw = @(Get-GraphProp -Object $d -Name 'registeredOwners')
            $ownerNames = [System.Collections.Generic.List[string]]::new()
            $ownerUpns = [System.Collections.Generic.List[string]]::new()

            foreach ($owner in $ownersRaw) {
                if ($null -eq $owner) { continue }
                $ownerId = [string](Get-GraphProp -Object $owner -Name 'id')
                $ownerName = [string](Get-GraphProp -Object $owner -Name 'displayName')
                $ownerUpn = [string](Get-GraphProp -Object $owner -Name 'userPrincipalName')
                if ($ownerId) { [void]$resolvedOwnerIds.Add($ownerId) }
                if ($ownerName) { $ownerNames.Add($ownerName) }
                if ($ownerUpn) { $ownerUpns.Add($ownerUpn) }
            }

            $isCompliant = Get-GraphProp -Object $d -Name 'isCompliant'
            $compliance = if ($null -eq $isCompliant) {
                'Unknown'
            }
            elseif ($isCompliant -eq $true -or "$isCompliant" -eq 'True') {
                'Compliant'
            }
            else {
                'NonCompliant'
            }

            if ($compliance -eq 'Compliant') { $compliantCount++ }
            elseif ($compliance -eq 'NonCompliant') { $nonCompliantCount++ }
            else { $unknownComplianceCount++ }

            if ($ownerNames.Count -eq 0) { $noOwnerCount++ }

            $deviceRows.Add([PSCustomObject]@{
                    DisplayName        = Get-GraphProp -Object $d -Name 'displayName'
                    DeviceId           = Get-GraphProp -Object $d -Name 'deviceId'
                    ObjectId           = Get-GraphProp -Object $d -Name 'id'
                    OperatingSystem    = Get-GraphProp -Object $d -Name 'operatingSystem'
                    AccountEnabled     = Get-GraphProp -Object $d -Name 'accountEnabled'
                    IsCompliant        = $isCompliant
                    ComplianceStatus   = $compliance
                    RegisteredOwners   = ($ownerNames -join '; ')
                    RegisteredOwnerUPN = ($ownerUpns -join '; ')
                })
        }

        Write-RunbookLog "  Devices page $pageNumber : +$($pageDevices.Count) (total $($deviceRows.Count))"
        $pageUri = Get-GraphNextLink -Page $page
        if ($pageUri) {
            Start-Sleep -Milliseconds 200
        }
    }

    Write-RunbookLog "Retrieved $($deviceRows.Count) device(s); $($resolvedOwnerIds.Count) unique owner id(s)."

    $summaryRows = @(
        [PSCustomObject]@{ Metric = 'TotalDevices'; Value = $deviceRows.Count }
        [PSCustomObject]@{ Metric = 'Compliant'; Value = $compliantCount }
        [PSCustomObject]@{ Metric = 'NonCompliant'; Value = $nonCompliantCount }
        [PSCustomObject]@{ Metric = 'ComplianceUnknown'; Value = $unknownComplianceCount }
        [PSCustomObject]@{ Metric = 'NoRegisteredOwner'; Value = $noOwnerCount }
        [PSCustomObject]@{ Metric = 'ResolvedOwners'; Value = $resolvedOwnerIds.Count }
    )

    @($deviceRows.ToArray()) | Export-Csv -LiteralPath $devicesPath -NoTypeInformation -Encoding UTF8
    @($summaryRows) | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

    Publish-ReportToSharePoint `
        -FilePaths @($devicesPath, $summaryPath) `
        -SiteUrl $SharePointSiteUrl `
        -FolderPath $SharePointFolderPath

    $duration = (Get-Date) - $runStart
    Write-RunbookLog 'Report complete.'
    Write-RunbookLog "  Devices : $($deviceRows.Count) -> $devicesPath"
    Write-RunbookLog "  Summary : $($summaryRows.Count) -> $summaryPath"
    Write-RunbookLog "  SharePoint : $SharePointSiteUrl / $SharePointFolderPath"
    Write-RunbookLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)
}
catch {
    $step = if ($_.InvocationInfo.Line) { $_.InvocationInfo.Line.Trim() } else { 'unknown' }
    Write-RunbookLog "ERROR: $($_.Exception.Message)"
    Write-RunbookLog "Failed near: $step"
    if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges') {
        Write-RunbookLog 'Missing Graph app permission on the Managed Identity. Ensure Device.Read.All (and ideally Directory.Read.All) + User.Read.All are assigned, then wait for token refresh.'
    }
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
