<#
.SYNOPSIS
    Exports a complete Microsoft 365 licensing report (Azure Automation runbook).

.DESCRIPTION
    Produces four CSV files:
      - TenantSummary  : purchased, consumed, and available units per SKU
      - GroupLicenses  : licenses assigned to each Entra group (group-based licensing)
      - Users          : one row per user with identity, sign-in, and license summary
      - LicenseDetail  : one row per user/license with assigning group and service plans

    Designed for Azure Automation with a **system-assigned Managed Identity**.
    CSVs are written under $env:TEMP, then uploaded (and overwritten) into a
    SharePoint document library folder via Connect-PnPOnline -ManagedIdentity.

.NOTES
    Azure Automation setup
    ----------------------
    1. Enable the **system-assigned** Managed Identity on the Automation Account
       (Identity > System assigned > On).
    2. Grant the identity these Microsoft Graph APPLICATION permissions
       (Entra ID > Enterprise applications > your MI > Permissions, or use
       App role assignments via Graph/PowerShell), then grant admin consent:

       Permission              | Used by
       ------------------------|--------------------------------------------------
       User.Read.All           | Users, managers, sign-in activity, license detail
       Group.Read.All          | Groups with assigned licenses (group licensing)
       GroupMember.Read.All    | Member counts on licensing groups
       AuditLog.Read.All       | Sign-in activity (interactive / non-interactive)
       Organization.Read.All   | Tenant SKU summary (Get-MgSubscribedSku)
       Sites.ReadWrite.All     | Upload CSVs via PnP (or Sites.Selected + site grant)
                                 Also grant the MI access to the target SharePoint site.

    3. Import these modules into the Automation Account (PowerShell 7.2 runtime).
       ALL Microsoft.Graph.* modules below MUST be the SAME version (e.g. all 2.38.1):
       - Microsoft.Graph.Authentication
       - Microsoft.Graph.Users
       - Microsoft.Graph.Groups
       - Microsoft.Graph.Identity.DirectoryManagement
       - PnP.PowerShell

    4. Set Automation Variables:
       SHAREPOINT_SITE_URL     e.g. https://contoso.sharepoint.com/sites/IT
       SHAREPOINT_FOLDER_PATH  e.g. Shared Documents/LicensingReports
                               Defaults to Shared Documents/LicensingReports if unset.

    Sign-in activity requires Entra ID P1 or P2 for data to be populated.

.EXAMPLE
    .\Licensing.ps1
#>
[CmdletBinding()]
Param()

$ErrorActionPreference = 'Stop'
$BatchSize = 10
$BatchDelayMs = 1000
$GraphMaxRetries = 8
$GraphBaseDelaySeconds = 10

# SharePoint destination from Automation Variables
$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
$SharePointFolderPath = Get-AutomationVariable -Name 'SHAREPOINT_FOLDER_PATH' -ErrorAction SilentlyContinue

if ([string]::IsNullOrWhiteSpace($SharePointFolderPath)) {
    $SharePointFolderPath = 'Shared Documents/LicensingReports'
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
        'Microsoft.Graph.Groups'
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
  - Microsoft.Graph.Groups
  - Microsoft.Graph.Identity.DirectoryManagement

Currently available:
$($installed -join [Environment]::NewLine)

Fix: run AzureAutomation\Update-GraphModules.ps1 (set your Automation Account name),
or in Portal > Automation Account > Modules update Users and Groups to match Authentication ($($authVersions[0])).
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
        [int]$Size = 20,
        [int]$DelayMs = 200,
        [string]$Activity = 'Graph batch'
    )

    $allResponses = [System.Collections.Generic.List[object]]::new()
    $totalBatches = [Math]::Ceiling($AllRequests.Count / [double]$Size)

    for ($i = 0; $i -lt $AllRequests.Count; $i += $Size) {
        $batchNum = [int]($i / $Size) + 1
        $end = [Math]::Min($i + $Size - 1, $AllRequests.Count - 1)
        $chunk = @($AllRequests[$i..$end])

        Write-RunbookLog "$Activity - batch $batchNum of $totalBatches"
        $batchResponses = Invoke-MgGraphBatch -Requests $chunk -MaxRetries $GraphMaxRetries

        # Retry individual throttled items inside a successful batch payload
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

    return $allResponses
}

function Import-GroupCacheBatch {
    param(
        [string[]]$GroupIds,
        [hashtable]$GroupCache,
        [int]$Size,
        [int]$DelayMs
    )

    $missing = @($GroupIds | Where-Object { $_ -and -not $GroupCache.ContainsKey($_) } | Select-Object -Unique)
    if ($missing.Count -eq 0) { return }

    Write-RunbookLog "Resolving $($missing.Count) assigning group(s)..."
    $requests = for ($i = 0; $i -lt $missing.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/groups/$($missing[$i])?`$select=id,displayName,mailNickname" }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Resolving groups'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $groupId = $missing[$index]
        if ($response.status -eq 200) {
            $GroupCache[$groupId] = [PSCustomObject]@{
                Id           = $response.body.id
                DisplayName  = $response.body.displayName
                MailNickname = $response.body.mailNickname
            }
        }
        else {
            $GroupCache[$groupId] = [PSCustomObject]@{
                Id           = $groupId
                DisplayName  = $groupId
                MailNickname = ''
            }
        }
    }
}

function Import-UserManagersBatch {
    param(
        [string[]]$UserIds,
        [hashtable]$ManagerCache,
        [int]$Size,
        [int]$DelayMs
    )

    $pending = @($UserIds | Where-Object { $_ -and -not $ManagerCache.ContainsKey($_) })
    if ($pending.Count -eq 0) { return }

    Write-RunbookLog "Loading managers for $($pending.Count) user(s) via batch..."
    $requests = for ($i = 0; $i -lt $pending.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/users/$($pending[$i])/manager?`$select=id,displayName,userPrincipalName" }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Loading managers'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $userId = $pending[$index]
        if ($response.status -eq 200 -and $response.body.id) {
            $ManagerCache[$userId] = [PSCustomObject]@{
                Name = $response.body.displayName
                UPN  = $response.body.userPrincipalName
            }
        }
        else {
            $ManagerCache[$userId] = [PSCustomObject]@{ Name = ''; UPN = '' }
        }
    }
}

function Get-SkuInfoFromCatalog {
    param(
        [string]$SkuId,
        [hashtable]$SkuCatalog
    )

    $partNumber = $SkuCatalog[$SkuId]
    if (-not $partNumber) { $partNumber = $SkuId }
    return @{
        SkuPartNumber = $partNumber
        FriendlyName  = Get-LicenseFriendlyName -SkuPartNumber $partNumber
    }
}

function Get-LicenseAssignmentForSku {
    param(
        $LicenseAssignmentStates,
        [string]$SkuId,
        [hashtable]$GroupCache
    )

    $statesForSku = @($LicenseAssignmentStates | Where-Object { $_.SkuId.ToString() -eq $SkuId })
    $state = $statesForSku | Select-Object -First 1
    $groupNames = [System.Collections.Generic.List[string]]::new()
    $groupIds = [System.Collections.Generic.List[string]]::new()
    $hasDirect = $false

    foreach ($skuState in $statesForSku) {
        $groupInfo = Resolve-AssigningGroup -GroupId $skuState.AssignedByGroup -GroupCache $GroupCache
        if ($groupInfo.Method -eq 'Group') {
            $groupNames.Add($groupInfo.Name)
            $groupIds.Add($groupInfo.Id)
        }
        else {
            $hasDirect = $true
        }
    }

    $method = if ($groupNames.Count -gt 0 -and $hasDirect) { 'Mixed' }
    elseif ($groupNames.Count -gt 0) { 'Group' }
    elseif ($hasDirect) { 'Direct' }
    else { 'Unknown' }

    return [PSCustomObject]@{
        AssignmentMethod = $method
        GroupNames       = $groupNames
        GroupIds         = $groupIds
        AssignmentState  = if ($state) { $state.State } else { 'Unknown' }
        AssignmentError  = ($statesForSku | Where-Object { $_.Error } | ForEach-Object { $_.Error }) -join '; '
    }
}

function Get-LicenseFriendlyName {
    param([string]$SkuPartNumber)

    $names = @{
        'ENTERPRISEPACK'           = 'Microsoft 365 E3'
        'ENTERPRISEPREMIUM'        = 'Microsoft 365 E5'
        'SPE_E3'                   = 'Microsoft 365 E3'
        'SPE_E5'                   = 'Microsoft 365 E5'
        'SPE_F1'                   = 'Microsoft 365 F3'
        'M365_F1'                  = 'Microsoft 365 F1'
        'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
        'O365_BUSINESS_PREMIUM'    = 'Microsoft 365 Business Standard'
        'SPB'                      = 'Microsoft 365 Business Premium'
        'SMB_BUSINESS_PREMIUM'     = 'Microsoft 365 Business Premium'
        'EXCHANGESTANDARD'         = 'Exchange Online Plan 1'
        'EXCHANGEENTERPRISE'       = 'Exchange Online Plan 2'
        'SHAREPOINTSTANDARD'       = 'SharePoint Online Plan 1'
        'SHAREPOINTENTERPRISE'     = 'SharePoint Online Plan 2'
        'PROJECTPROFESSIONAL'      = 'Project Plan 3'
        'PROJECTPREMIUM'           = 'Project Plan 5'
        'VISIOCLIENT'              = 'Visio Plan 2'
        'POWER_BI_PRO'             = 'Power BI Pro'
        'POWER_BI_STANDARD'        = 'Power BI Free'
        'FLOW_FREE'                = 'Power Automate Free'
        'TEAMS_EXPLORATORY'        = 'Microsoft Teams Exploratory'
        'MCOPSTN2'                 = 'Microsoft Teams Phone Standard'
        'MCOEV'                    = 'Microsoft Teams Phone'
        'Microsoft_365_Copilot'    = 'Microsoft 365 Copilot'
        'WINDOWS_STORE'            = 'Windows Store for Business'
        'RIGHTSMANAGEMENT'         = 'Azure Information Protection Plan 1'
        'EMS'                      = 'Enterprise Mobility + Security E3'
        'EMSPREMIUM'               = 'Enterprise Mobility + Security E5'
        'AAD_PREMIUM'              = 'Microsoft Entra ID P1'
        'AAD_PREMIUM_P2'           = 'Microsoft Entra ID P2'
        'ATP_ENTERPRISE'           = 'Microsoft Defender for Office 365 Plan 2'
        'THREAT_INTELLIGENCE'      = 'Microsoft Defender for Office 365 Plan 2'
        'DEVELOPERPACK_E5'         = 'Microsoft 365 E5 Developer'
    }

    if ($names.ContainsKey($SkuPartNumber)) {
        return $names[$SkuPartNumber]
    }
    return $SkuPartNumber
}

function Get-InactiveDays {
    param($DateTime)

    if (-not $DateTime) { return $null }
    return [int](New-TimeSpan -Start $DateTime -End (Get-Date)).TotalDays
}

function Get-SignInActivityInfo {
    param($SignInActivity)

    if ($null -eq $SignInActivity) {
        return [PSCustomObject]@{
            LastInteractiveSignInDateTime    = $null
            LastNonInteractiveSignInDateTime = $null
            InactiveDaysInteractiveSignIn    = $null
            InactiveDaysNonInteractiveSignIn = $null
        }
    }

    $interactive = $null
    $nonInteractive = $null

    if ($SignInActivity -is [System.Collections.IDictionary]) {
        $interactive = $SignInActivity['lastSignInDateTime']
        if (-not $interactive) { $interactive = $SignInActivity['LastSignInDateTime'] }
        $nonInteractive = $SignInActivity['lastNonInteractiveSignInDateTime']
        if (-not $nonInteractive) { $nonInteractive = $SignInActivity['LastNonInteractiveSignInDateTime'] }
    }
    else {
        $interactive = $SignInActivity.lastSignInDateTime
        if (-not $interactive) { $interactive = $SignInActivity.LastSignInDateTime }
        $nonInteractive = $SignInActivity.lastNonInteractiveSignInDateTime
        if (-not $nonInteractive) { $nonInteractive = $SignInActivity.LastNonInteractiveSignInDateTime }
    }

    return [PSCustomObject]@{
        LastInteractiveSignInDateTime    = $interactive
        LastNonInteractiveSignInDateTime = $nonInteractive
        InactiveDaysInteractiveSignIn    = Get-InactiveDays -DateTime $interactive
        InactiveDaysNonInteractiveSignIn = Get-InactiveDays -DateTime $nonInteractive
    }
}

function Resolve-AssigningGroup {
    param(
        $GroupId,
        [hashtable]$GroupCache
    )

    if (-not $GroupId) {
        return @{
            Method = 'Direct'
            Name   = ''
            Id     = ''
        }
    }

    $id = $GroupId.ToString()
    if (-not $GroupCache.ContainsKey($id)) {
        try {
            $group = Invoke-WithGraphRetry -Activity "Get-MgGroup $id" -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
                Get-MgGroup -GroupId $id -Property Id, DisplayName, MailNickname
            }
            $GroupCache[$id] = $group
        }
        catch {
            $GroupCache[$id] = [PSCustomObject]@{
                Id           = $id
                DisplayName  = $id
                MailNickname = ''
            }
        }
    }

    return @{
        Method = 'Group'
        Name   = $GroupCache[$id].DisplayName
        Id     = $id
    }
}

function Get-UserManagerInfo {
    param(
        [string]$UserId,
        [hashtable]$ManagerCache
    )

    if ($ManagerCache.ContainsKey($UserId)) {
        return $ManagerCache[$UserId]
    }

    return [PSCustomObject]@{ Name = ''; UPN = '' }
}

function Get-AssignmentStateSummary {
    param(
        $LicenseAssignmentStates,
        [hashtable]$SkuCatalog,
        [hashtable]$GroupCache
    )

    if (-not $LicenseAssignmentStates -or @($LicenseAssignmentStates).Count -eq 0) {
        return '', '', 'None', @()
    }

    $licensingGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $hasDirect = $false
    $states = foreach ($state in $LicenseAssignmentStates) {
        $skuName = $SkuCatalog[$state.SkuId.ToString()]
        if (-not $skuName) { $skuName = $state.SkuId }

        $groupInfo = Resolve-AssigningGroup -GroupId $state.AssignedByGroup -GroupCache $GroupCache
        if ($groupInfo.Method -eq 'Group') {
            [void]$licensingGroups.Add($groupInfo.Name)
            $assignedBy = "Group:$($groupInfo.Name)"
        }
        else {
            $hasDirect = $true
            $assignedBy = 'Direct'
        }

        "$skuName=$($state.State)($assignedBy)"
    }

    $errors = ($LicenseAssignmentStates | Where-Object { $_.State -eq 'ActiveWithError' -and $_.Error } |
        ForEach-Object { "$($SkuCatalog[$_.SkuId.ToString()]):$($_.Error)" }) -join '; '

    $method = if ($licensingGroups.Count -gt 0 -and $hasDirect) {
        'Mixed'
    }
    elseif ($licensingGroups.Count -gt 0) {
        'Group'
    }
    elseif ($hasDirect) {
        'Direct'
    }
    else {
        'None'
    }

    return ($states -join '; '), $errors, $method, @($licensingGroups)
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

function Publish-ReportToSharePoint {
    param(
        [Parameter(Mandatory)]
        [string[]]$FilePaths,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [string]$FolderPath = 'Shared Documents/LicensingReports'
    )

    Import-Module PnP.PowerShell -ErrorAction Stop

    $folder = Get-PnPDocumentFolderPath -FolderPath $FolderPath
    Write-RunbookLog "Uploading reports to SharePoint site '$SiteUrl' (folder: '$folder') via Connect-PnPOnline..."

    Connect-PnPOnline -Url $SiteUrl -ManagedIdentity

    try {
        foreach ($filePath in $FilePaths) {
            if (-not (Test-Path -LiteralPath $filePath)) {
                throw "Report file not found: $filePath"
            }

            $fileName = Split-Path -Path $filePath -Leaf
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

Connect-MgGraphManagedIdentity

$tenantSummaryPath = Join-Path $env:TEMP 'M365-LicensingReport-TenantSummary.csv'
$groupLicensesPath = Join-Path $env:TEMP 'M365-LicensingReport-GroupLicenses.csv'
$userSummaryPath = Join-Path $env:TEMP 'M365-LicensingReport-Users.csv'
$licenseDetailPath = Join-Path $env:TEMP 'M365-LicensingReport-LicenseDetail.csv'

$groupCache = @{}
$managerCache = @{}

Write-RunbookLog 'Building tenant license catalog...'
$subscribedSkus = @(Invoke-WithGraphRetry -Activity 'Get-MgSubscribedSku' -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
        Get-MgSubscribedSku -All
    })
$skuCatalog = @{}
$tenantSummary = [System.Collections.Generic.List[object]]::new()

foreach ($sku in $subscribedSkus) {
    $purchased = [int]$sku.PrepaidUnits.Enabled
    $consumed = [int]$sku.ConsumedUnits
    $available = [Math]::Max(0, $purchased - $consumed)
    $friendlyName = Get-LicenseFriendlyName -SkuPartNumber $sku.SkuPartNumber

    $skuCatalog[$sku.SkuId.ToString()] = $sku.SkuPartNumber

    $tenantSummary.Add([PSCustomObject]@{
            SkuPartNumber     = $sku.SkuPartNumber
            FriendlyName      = $friendlyName
            SkuId             = $sku.SkuId
            PurchasedLicenses = $purchased
            ConsumedLicenses  = $consumed
            AvailableLicenses = $available
            CapabilityStatus  = $sku.CapabilityStatus
            AppliesTo         = $sku.AppliesTo
        })
}

Write-RunbookLog 'Retrieving groups with assigned licenses...'
$licensedGroups = @(Invoke-WithGraphRetry -Activity 'Get-MgGroup' -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
        Get-MgGroup -All -Property Id, DisplayName, MailNickname, AssignedLicenses, LicenseProcessingState |
            Where-Object { @($_.AssignedLicenses).Count -gt 0 }
    })
$groupLicenses = [System.Collections.Generic.List[object]]::new()
$groupIndex = 0

foreach ($group in $licensedGroups) {
    $groupIndex++
    $groupCache[$group.Id] = $group

    $groupSkuPartNumbers = [System.Collections.Generic.List[string]]::new()
    $groupFriendlyNames = [System.Collections.Generic.List[string]]::new()

    foreach ($assignedLicense in $group.AssignedLicenses) {
        $partNumber = $skuCatalog[$assignedLicense.SkuId.ToString()]
        if (-not $partNumber) { $partNumber = $assignedLicense.SkuId.ToString() }
        $groupSkuPartNumbers.Add($partNumber)
        $groupFriendlyNames.Add((Get-LicenseFriendlyName -SkuPartNumber $partNumber))
    }

    $memberCount = $null
    try {
        $memberCount = @(Invoke-WithGraphRetry -Activity "Get-MgGroupMember $($group.DisplayName)" -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
                Get-MgGroupMember -GroupId $group.Id -All -Property Id
            }).Count
    }
    catch {
        Write-RunbookLog "WARNING: Could not get member count for group '$($group.DisplayName)': $($_.Exception.Message)"
    }

    if ($groupIndex -lt $licensedGroups.Count) {
        Start-Sleep -Milliseconds 250
    }

    $groupLicenses.Add([PSCustomObject]@{
            GroupDisplayName       = $group.DisplayName
            GroupMailNickname      = $group.MailNickname
            GroupId                = $group.Id
            AssignedLicenses       = $groupFriendlyNames -join '; '
            AssignedSkuPartNumbers = $groupSkuPartNumbers -join '; '
            MemberCount            = $memberCount
            LicenseProcessingState = $group.LicenseProcessingState
        })
}

Write-RunbookLog "Found $($licensedGroups.Count) licensing group(s)."

Write-RunbookLog 'Retrieving all tenant users with managers and sign-in activity (paged)...'
$users = [System.Collections.Generic.List[object]]::new()
# $expand=manager loads managers in the same pages (avoids ~1200 separate manager batch calls)
# signInActivity requires AuditLog.Read.All and Entra ID P1/P2
$pageUri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,mail,userType,accountEnabled,department,jobTitle,companyName,createdDateTime,assignedLicenses,licenseAssignmentStates,signInActivity&$expand=manager($select=id,displayName,userPrincipalName)&$top=100'
$pageNumber = 0

while (-not [string]::IsNullOrWhiteSpace($pageUri)) {
    $pageNumber++
    $page = Invoke-WithGraphRetry -Activity "Get users page $pageNumber" -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri $pageUri
    }

    $pageUsers = if ($page -is [System.Collections.IDictionary]) {
        @($page['value'])
    }
    else {
        @($page.value)
    }

    foreach ($u in $pageUsers) {
        if ($null -eq $u) { continue }

        $assignedLicenses = @(
            foreach ($lic in @($u.assignedLicenses)) {
                if ($null -eq $lic) { continue }
                [PSCustomObject]@{
                    SkuId         = $lic.skuId
                    DisabledPlans = $lic.disabledPlans
                }
            }
        )

        $licenseAssignmentStates = @(
            foreach ($st in @($u.licenseAssignmentStates)) {
                if ($null -eq $st) { continue }
                [PSCustomObject]@{
                    SkuId           = $st.skuId
                    AssignedByGroup = $st.assignedByGroup
                    State           = $st.state
                    Error           = $st.error
                }
            }
        )

        $managerObj = $u.manager
        if ($page -is [System.Collections.IDictionary] -or $u -is [System.Collections.IDictionary]) {
            if ($u -is [System.Collections.IDictionary]) {
                $managerObj = $u['manager']
            }
        }

        $managerName = ''
        $managerUpn = ''
        if ($null -ne $managerObj) {
            if ($managerObj -is [System.Collections.IDictionary]) {
                $managerName = [string]$managerObj['displayName']
                $managerUpn = [string]$managerObj['userPrincipalName']
            }
            else {
                $managerName = [string]$managerObj.displayName
                $managerUpn = [string]$managerObj.userPrincipalName
            }
        }

        $userId = if ($u -is [System.Collections.IDictionary]) { [string]$u['id'] } else { [string]$u.id }
        $managerCache[$userId] = [PSCustomObject]@{
            Name = $managerName
            UPN  = $managerUpn
        }

        $signInActivity = if ($u -is [System.Collections.IDictionary]) { $u['signInActivity'] } else { $u.signInActivity }
        if ($null -ne $signInActivity) {
            if ($signInActivity -is [System.Collections.IDictionary]) {
                $signInActivity = [PSCustomObject]@{
                    LastSignInDateTime               = $signInActivity['lastSignInDateTime']
                    LastNonInteractiveSignInDateTime = $signInActivity['lastNonInteractiveSignInDateTime']
                }
            }
            else {
                $signInActivity = [PSCustomObject]@{
                    LastSignInDateTime               = $signInActivity.lastSignInDateTime
                    LastNonInteractiveSignInDateTime = $signInActivity.lastNonInteractiveSignInDateTime
                }
            }
        }

        $users.Add([PSCustomObject]@{
                Id                      = $userId
                UserPrincipalName       = if ($u -is [System.Collections.IDictionary]) { $u['userPrincipalName'] } else { $u.userPrincipalName }
                DisplayName             = if ($u -is [System.Collections.IDictionary]) { $u['displayName'] } else { $u.displayName }
                Mail                    = if ($u -is [System.Collections.IDictionary]) { $u['mail'] } else { $u.mail }
                UserType                = if ($u -is [System.Collections.IDictionary]) { $u['userType'] } else { $u.userType }
                AccountEnabled          = if ($u -is [System.Collections.IDictionary]) { $u['accountEnabled'] } else { $u.accountEnabled }
                Department              = if ($u -is [System.Collections.IDictionary]) { $u['department'] } else { $u.department }
                JobTitle                = if ($u -is [System.Collections.IDictionary]) { $u['jobTitle'] } else { $u.jobTitle }
                CompanyName             = if ($u -is [System.Collections.IDictionary]) { $u['companyName'] } else { $u.companyName }
                CreatedDateTime         = if ($u -is [System.Collections.IDictionary]) { $u['createdDateTime'] } else { $u.createdDateTime }
                AssignedLicenses        = $assignedLicenses
                LicenseAssignmentStates = $licenseAssignmentStates
                SignInActivity          = $signInActivity
            })
    }

    Write-RunbookLog "  Users page $pageNumber : +$($pageUsers.Count) (total $($users.Count))"

    if ($page -is [System.Collections.IDictionary]) {
        $pageUri = $page['@odata.nextLink']
    }
    else {
        $pageUri = $page.'@odata.nextLink'
    }

    if ($pageUri) {
        Start-Sleep -Milliseconds 200
    }
}

$users = @($users)
Write-RunbookLog "Retrieved $($users.Count) user(s) with manager data."

$assigningGroupIds = @(
    $users | ForEach-Object { $_.LicenseAssignmentStates } |
    Where-Object { $_.AssignedByGroup } |
    ForEach-Object { $_.AssignedByGroup.ToString() } |
    Select-Object -Unique
)
Import-GroupCacheBatch -GroupIds $assigningGroupIds -GroupCache $groupCache -Size $BatchSize -DelayMs $BatchDelayMs

Write-RunbookLog 'Building report rows...'
$userSummary = [System.Collections.Generic.List[object]]::new()
$licenseDetail = [System.Collections.Generic.List[object]]::new()
$processed = 0

foreach ($user in $users) {
    $processed++
    if ($processed % 500 -eq 0 -or $processed -eq $users.Count) {
        Write-RunbookLog "Building report: $processed / $($users.Count)"
    }

    $upn = $user.UserPrincipalName
    $licenseCount = @($user.AssignedLicenses).Count
    $accountStatus = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }

    $manager = Get-UserManagerInfo -UserId $user.Id -ManagerCache $managerCache
    $signIn = Get-SignInActivityInfo -SignInActivity $user.SignInActivity

    $assignmentSummary, $assignmentErrors, $assignmentMethod, $licensingGroups = Get-AssignmentStateSummary `
        -LicenseAssignmentStates $user.LicenseAssignmentStates `
        -SkuCatalog $skuCatalog `
        -GroupCache $groupCache

    $skuPartNumbers = [System.Collections.Generic.List[string]]::new()
    $friendlyNames = [System.Collections.Generic.List[string]]::new()

    if ($licenseCount -gt 0) {
        foreach ($assigned in $user.AssignedLicenses) {
            $skuId = $assigned.SkuId.ToString()
            $skuInfo = Get-SkuInfoFromCatalog -SkuId $skuId -SkuCatalog $skuCatalog
            $partNumber = $skuInfo.SkuPartNumber
            $friendlyName = $skuInfo.FriendlyName
            $skuPartNumbers.Add($partNumber)
            $friendlyNames.Add($friendlyName)

            $assignment = Get-LicenseAssignmentForSku `
                -LicenseAssignmentStates $user.LicenseAssignmentStates `
                -SkuId $skuId `
                -GroupCache $groupCache

            $licenseDetail.Add([PSCustomObject]@{
                    UserPrincipalName                = $upn
                    DisplayName                      = $user.DisplayName
                    Manager                          = $manager.Name
                    ManagerUPN                       = $manager.UPN
                    LastInteractiveSignInDateTime    = $signIn.LastInteractiveSignInDateTime
                    LastNonInteractiveSignInDateTime = $signIn.LastNonInteractiveSignInDateTime
                    InactiveDaysInteractiveSignIn    = $signIn.InactiveDaysInteractiveSignIn
                    InactiveDaysNonInteractiveSignIn = $signIn.InactiveDaysNonInteractiveSignIn
                    SkuPartNumber                    = $partNumber
                    FriendlyName                     = $friendlyName
                    SkuId                            = $assigned.SkuId
                    AssignmentMethod                 = $assignment.AssignmentMethod
                    AssigningGroupName               = ($assignment.GroupNames -join '; ')
                    AssigningGroupId                 = ($assignment.GroupIds -join '; ')
                    AssignmentState                  = $assignment.AssignmentState
                    AssignmentError                  = $assignment.AssignmentError
                })
        }
    }

    $userSummary.Add([PSCustomObject]@{
            UserPrincipalName                = $upn
            DisplayName                      = $user.DisplayName
            Mail                             = $user.Mail
            Manager                          = $manager.Name
            ManagerUPN                       = $manager.UPN
            UserType                         = $user.UserType
            AccountStatus                    = $accountStatus
            Department                       = $user.Department
            JobTitle                         = $user.JobTitle
            CompanyName                      = $user.CompanyName
            CreatedDateTime                  = $user.CreatedDateTime
            LicenseCount                     = $licenseCount
            AssignedLicenses                 = if ($friendlyNames.Count -gt 0) { $friendlyNames -join '; ' } else { 'No License Assigned' }
            AssignedSkuPartNumbers           = if ($skuPartNumbers.Count -gt 0) { $skuPartNumbers -join '; ' } else { '' }
            LicenseAssignmentMethod          = $assignmentMethod
            LicensingGroups                  = ($licensingGroups -join '; ')
            LicenseAssignmentStates          = $assignmentSummary
            LicenseAssignmentErrors          = $assignmentErrors
            LastInteractiveSignInDateTime    = $signIn.LastInteractiveSignInDateTime
            LastNonInteractiveSignInDateTime = $signIn.LastNonInteractiveSignInDateTime
            InactiveDaysInteractiveSignIn    = $signIn.InactiveDaysInteractiveSignIn
            InactiveDaysNonInteractiveSignIn = $signIn.InactiveDaysNonInteractiveSignIn
            ObjectId                         = $user.Id
        })
}

@($tenantSummary.ToArray()) | Export-Csv -LiteralPath $tenantSummaryPath -NoTypeInformation -Encoding UTF8
@($groupLicenses.ToArray()) | Export-Csv -LiteralPath $groupLicensesPath -NoTypeInformation -Encoding UTF8
@($userSummary.ToArray()) | Export-Csv -LiteralPath $userSummaryPath -NoTypeInformation -Encoding UTF8
@($licenseDetail.ToArray()) | Export-Csv -LiteralPath $licenseDetailPath -NoTypeInformation -Encoding UTF8

$reportFiles = @(
    $tenantSummaryPath
    $groupLicensesPath
    $userSummaryPath
    $licenseDetailPath
)

Publish-ReportToSharePoint `
    -FilePaths $reportFiles `
    -SiteUrl $SharePointSiteUrl `
    -FolderPath $SharePointFolderPath

Write-RunbookLog 'Report complete:'
Write-RunbookLog "  Tenant SKUs    : $($tenantSummary.Count) rows -> $tenantSummaryPath"
Write-RunbookLog "  Group licenses : $($groupLicenses.Count) rows -> $groupLicensesPath"
Write-RunbookLog "  Users          : $($userSummary.Count) rows -> $userSummaryPath"
Write-RunbookLog "  License detail : $($licenseDetail.Count) rows -> $licenseDetailPath"
Write-RunbookLog "  SharePoint     : $SharePointSiteUrl / $SharePointFolderPath"

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
