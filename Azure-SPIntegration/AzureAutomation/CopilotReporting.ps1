<#
.SYNOPSIS
    Exports Microsoft 365 Copilot usage user detail (Azure Automation).

.DESCRIPTION
    Calls Microsoft Graph getMicrosoft365CopilotUsageUserDetail for the configured
    period, exports user-detail + summary CSVs, then uploads them to SharePoint
    Documents via Connect-PnPOnline -ManagedIdentity.

    Graph returns either JSON (preferred) or a 302 redirect to a temporary CSV
    download URL when $format=text/csv. This runbook uses JSON and writes CSV
    locally so SharePoint always receives stable, overwriteable file names.

.NOTES
    Azure Automation setup
    ----------------------
    1. Enable system-assigned Managed Identity on the Automation Account.
    2. Grant Graph application permissions (admin consent):
         Reports.Read.All
         Sites.ReadWrite.All   (or SharePoint Sites.FullControl.All for PnP)
       Re-run Set-SystemManagedId.ps1 after updating permissions, then wait a few minutes.
    3. Runtime modules:
         Microsoft.Graph.Authentication
         PnP.PowerShell
    4. Automation Variables:
         SHAREPOINT_SITE_URL
         SHAREPOINT_FOLDER_PATH_COPILOT  (optional; default Shared Documents/CopilotReports)
         COPILOT_REPORT_PERIOD           (optional; D7|D30|D90|D180|ALL; default D180)
         COPILOT_REPORT_VERSION          (optional; v1|v2; default v2)

    Defaults match the Microsoft 365 Admin Center Copilot usage details view
    (v2 metrics over 180 days). CSV headers use Admin Center column labels.
    v1 is still available if you need last-activity-only reporting.

.EXAMPLE
    .\CopilotReporting.ps1

.EXAMPLE
    .\CopilotReporting.ps1 -Period D30 -Version v1
#>
[CmdletBinding()]
Param(
    [ValidateSet('D7', 'D30', 'D90', 'D180', 'ALL')]
    [string]$Period = 'D180',

    [ValidateSet('v1', 'v2')]
    [string]$Version = 'v2',

    [string]$ExportPath = (Join-Path $env:TEMP 'Copilot-Report')
)

$ErrorActionPreference = 'Stop'
$GraphMaxRetries = 8
$GraphBaseDelaySeconds = 10
$runStart = Get-Date

$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
$SharePointFolderPath = Get-AutomationVariable -Name 'SHAREPOINT_FOLDER_PATH_COPILOT' -ErrorAction SilentlyContinue
$PeriodVariable = Get-AutomationVariable -Name 'COPILOT_REPORT_PERIOD' -ErrorAction SilentlyContinue
$VersionVariable = Get-AutomationVariable -Name 'COPILOT_REPORT_VERSION' -ErrorAction SilentlyContinue

if (-not [string]::IsNullOrWhiteSpace($PeriodVariable)) {
    $Period = $PeriodVariable.Trim().ToUpperInvariant()
}

if (-not [string]::IsNullOrWhiteSpace($VersionVariable)) {
    $Version = $VersionVariable.Trim().ToLowerInvariant()
}

function Write-RunbookLog {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Output "[$stamp] $Message"
}



if ($Period -notin @('D7', 'D30', 'D90', 'D180', 'ALL')) {
    throw "Invalid COPILOT_REPORT_PERIOD '$Period'. Use D7, D30, D90, D180, or ALL."
}

if ($Version -notin @('v1', 'v2')) {
    throw "Invalid COPILOT_REPORT_VERSION '$Version'. Use v1 or v2."
}

if ([string]::IsNullOrWhiteSpace($SharePointFolderPath)) {
    $SharePointFolderPath = 'Shared Documents/CopilotReports'
}

if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
    throw 'Automation variable SHAREPOINT_SITE_URL is required (e.g. https://contoso.sharepoint.com/sites/IT).'
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

function Connect-MgGraphManagedIdentity {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    Write-RunbookLog 'Connecting to Microsoft Graph with system-assigned Managed Identity...'
    Connect-MgGraph -Identity -NoWelcome

    if (Get-Command Set-MgRequestContext -ErrorAction SilentlyContinue) {
        Set-MgRequestContext -Retries $GraphMaxRetries -RetryDelay $GraphBaseDelaySeconds | Out-Null
    }

    $context = Get-MgContext
    Write-RunbookLog "Connected. AuthType=$($context.AuthType); TenantId=$($context.TenantId); AppName=$($context.AppName)"
    if ($context.Scopes) {
        Write-RunbookLog "Token roles/scopes: $($context.Scopes -join ', ')"
        if ($context.Scopes -notcontains 'Reports.Read.All') {
            Write-RunbookLog "WARNING: Token does not include Reports.Read.All. Confirm the Managed Identity has that app role, then wait 5-10 minutes."
        }
    }
    else {
        Write-RunbookLog 'WARNING: Get-MgContext.Scopes is empty; cannot verify Reports.Read.All in the access token.'
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

function Test-HasActivityDate {
    param($Value)
    -not [string]::IsNullOrWhiteSpace([string]$Value)
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
        'UsageSummary' { return 'Copilot-Report-UsageSummary.csv' }
        'UsageUserDetail' { return 'Copilot-Report-UsageUserDetail.csv' }
        default { return $leaf }
    }
}

function Publish-ReportToSharePoint {
    param(
        [Parameter(Mandatory)]
        [string[]]$FilePaths,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [string]$FolderPath = 'Shared Documents/CopilotReports'
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

function Get-CopilotReportPeriodValue {
    param($Item)

    $direct = Get-GraphProp -Object $Item -Name 'reportPeriod'
    if (-not [string]::IsNullOrWhiteSpace([string]$direct)) {
        return $direct
    }

    $byPeriod = @(Get-GraphProp -Object $Item -Name 'copilotActivityUserDetailsByPeriod')
    foreach ($entry in $byPeriod) {
        $periodValue = Get-GraphProp -Object $entry -Name 'reportPeriod'
        if (-not [string]::IsNullOrWhiteSpace([string]$periodValue)) {
            return $periodValue
        }
    }

    return $null
}

function Get-CopilotUsageUserDetailRows {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('D7', 'D30', 'D90', 'D180', 'ALL')]
        [string]$ReportPeriod,

        [Parameter(Mandatory)]
        [ValidateSet('v1', 'v2')]
        [string]$ReportVersion
    )

    # Prefer GA /copilot/reports (article + current docs). Fall back to legacy /reports URIs.
    $uris = @(
        "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod',version='$ReportVersion')?`$format=application/json"
        "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod')?`$format=application/json"
        "https://graph.microsoft.com/beta/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod',version='$ReportVersion')?`$format=application/json"
        "https://graph.microsoft.com/v1.0/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod')?`$format=application/json"
        "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod')?`$format=application/json"
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $pageUri = $null
    $page = $null
    $lastError = $null

    foreach ($candidate in $uris) {
        try {
            Write-RunbookLog "Requesting Copilot usage report: $candidate"
            $page = Invoke-WithGraphRetry -Activity 'Get Copilot usage user detail' -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
                Invoke-MgGraphRequest -Method GET -Uri $candidate
            }
            $pageUri = $candidate
            $lastError = $null
            break
        }
        catch {
            $lastError = $_
            if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges|AccessDenied') {
                throw
            }
            Write-RunbookLog "Endpoint unavailable, trying next: $($_.Exception.Message)"
        }
    }

    if (-not $pageUri) {
        throw "Unable to retrieve Copilot usage report. Last error: $($lastError.Exception.Message)"
    }

    $pageNumber = 0
    while ($true) {
        $pageNumber++
        if ($pageNumber -gt 1 -or $null -eq $page) {
            $page = Invoke-WithGraphRetry -Activity "Get Copilot usage page $pageNumber" -MaxRetries $GraphMaxRetries -BaseDelaySeconds $GraphBaseDelaySeconds -ScriptBlock {
                Invoke-MgGraphRequest -Method GET -Uri $pageUri
            }
        }

        $pageRows = Get-GraphCollectionValue -Page $page
        foreach ($item in $pageRows) {
            if ($null -eq $item) { continue }

            # Column names match Microsoft 365 Admin Center > Reports > Copilot usage details.
            $rows.Add([PSCustomObject]@{
                    'Username'                                      = Get-GraphProp -Object $item -Name 'userPrincipalName'
                    'Display name'                                  = Get-GraphProp -Object $item -Name 'displayName'
                    'Prompts submitted'                             = Get-GraphProp -Object $item -Name 'promptsSubmitted'
                    'Prompts submitted in Copilot Chat (work)'      = Get-GraphProp -Object $item -Name 'copilotChatWorkPromptsSubmitted'
                    'Prompts submitted in Copilot Chat (web)'       = Get-GraphProp -Object $item -Name 'copilotChatWebPromptsSubmitted'
                    'Active days'                                   = Get-GraphProp -Object $item -Name 'activeUsageDays'
                    'Last activity date (UTC)'                      = Get-GraphProp -Object $item -Name 'lastActivityDate'
                    'Last activity date of Copilot Chat'            = Get-GraphProp -Object $item -Name 'copilotChatLastActivityDate'
                    'Last activity date of Copilot Chat (work)'     = Get-GraphProp -Object $item -Name 'copilotChatWorkLastActivityDate'
                    'Last activity date of Copilot Chat (web)'      = Get-GraphProp -Object $item -Name 'copilotChatWebLastActivityDate'
                    'Last activity date of Microsoft Teams Copilot' = Get-GraphProp -Object $item -Name 'microsoftTeamsCopilotLastActivityDate'
                    'Last activity date of Word Copilot'            = Get-GraphProp -Object $item -Name 'wordCopilotLastActivityDate'
                    'Last activity date of Excel Copilot'           = Get-GraphProp -Object $item -Name 'excelCopilotLastActivityDate'
                    'Last activity date of PowerPoint Copilot'      = Get-GraphProp -Object $item -Name 'powerPointCopilotLastActivityDate'
                    'Last activity date of Outlook Copilot'         = Get-GraphProp -Object $item -Name 'outlookCopilotLastActivityDate'
                    'Last activity date of OneNote Copilot'         = Get-GraphProp -Object $item -Name 'oneNoteCopilotLastActivityDate'
                    'Last activity date of Loop Copilot'            = Get-GraphProp -Object $item -Name 'loopCopilotLastActivityDate'
                    'Last activity date of Microsoft 365 Copilot'   = Get-GraphProp -Object $item -Name 'microsoft365CopilotLastActivityDate'
                    'Last activity date of Edge Copilot'            = Get-GraphProp -Object $item -Name 'edgeCopilotLastActivityDate'
                    'Last activity date of Copilot Agent'           = Get-GraphProp -Object $item -Name 'copilotAgentLastActivityDate'
                    'Report refresh date'                           = Get-GraphProp -Object $item -Name 'reportRefreshDate'
                    'Report period'                                 = Get-CopilotReportPeriodValue -Item $item
                    'Report version'                                = $ReportVersion
                })
        }

        Write-RunbookLog "  Usage page $pageNumber : +$($pageRows.Count) (total $($rows.Count))"
        $pageUri = Get-GraphNextLink -Page $page
        if ([string]::IsNullOrWhiteSpace($pageUri)) {
            break
        }
        Start-Sleep -Milliseconds 200
        $page = $null
    }

    return @($rows.ToArray())
}

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------

try {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $exportRoot = Join-Path $env:TEMP ("Copilot-Report-" + $timestamp)
    if (-not (Test-Path -LiteralPath $exportRoot)) {
        New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
    }

    $detailPath = Join-Path $exportRoot "UsageUserDetail-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "UsageSummary-$timestamp.csv"

    Write-RunbookLog "Starting Microsoft 365 Copilot usage report (period=$Period; version=$Version)"
    Connect-MgGraphManagedIdentity

    $userRows = @(Get-CopilotUsageUserDetailRows -ReportPeriod $Period -ReportVersion $Version)

    $anyActivity = 0
    $chatActivity = 0
    $teamsActivity = 0
    $wordActivity = 0
    $excelActivity = 0
    $pptActivity = 0
    $outlookActivity = 0
    $oneNoteActivity = 0
    $loopActivity = 0
    $edgeActivity = 0
    $agentActivity = 0
    $totalPrompts = 0

    foreach ($row in $userRows) {
        if (Test-HasActivityDate $row.'Last activity date (UTC)') { $anyActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Copilot Chat') { $chatActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Microsoft Teams Copilot') { $teamsActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Word Copilot') { $wordActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Excel Copilot') { $excelActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of PowerPoint Copilot') { $pptActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Outlook Copilot') { $outlookActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of OneNote Copilot') { $oneNoteActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Loop Copilot') { $loopActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Edge Copilot') { $edgeActivity++ }
        if (Test-HasActivityDate $row.'Last activity date of Copilot Agent') { $agentActivity++ }

        $promptValue = 0
        if ([int]::TryParse([string]$row.'Prompts submitted', [ref]$promptValue)) {
            $totalPrompts += $promptValue
        }
    }

    $summaryRows = @(
        [PSCustomObject]@{ Metric = 'Report period'; Value = $Period }
        [PSCustomObject]@{ Metric = 'Report version'; Value = $Version }
        [PSCustomObject]@{ Metric = 'Total users in report'; Value = $userRows.Count }
        [PSCustomObject]@{ Metric = 'Users with any Copilot activity'; Value = $anyActivity }
        [PSCustomObject]@{ Metric = 'Users with Copilot Chat activity'; Value = $chatActivity }
        [PSCustomObject]@{ Metric = 'Users with Teams Copilot activity'; Value = $teamsActivity }
        [PSCustomObject]@{ Metric = 'Users with Word Copilot activity'; Value = $wordActivity }
        [PSCustomObject]@{ Metric = 'Users with Excel Copilot activity'; Value = $excelActivity }
        [PSCustomObject]@{ Metric = 'Users with PowerPoint Copilot activity'; Value = $pptActivity }
        [PSCustomObject]@{ Metric = 'Users with Outlook Copilot activity'; Value = $outlookActivity }
        [PSCustomObject]@{ Metric = 'Users with OneNote Copilot activity'; Value = $oneNoteActivity }
        [PSCustomObject]@{ Metric = 'Users with Loop Copilot activity'; Value = $loopActivity }
        [PSCustomObject]@{ Metric = 'Users with Edge Copilot activity'; Value = $edgeActivity }
        [PSCustomObject]@{ Metric = 'Users with Copilot Agent activity'; Value = $agentActivity }
        [PSCustomObject]@{ Metric = 'Total prompts submitted'; Value = $totalPrompts }
        [PSCustomObject]@{ Metric = 'Generated at (UTC)'; Value = (Get-Date).ToUniversalTime().ToString('o') }
    )

    @($userRows) | Export-Csv -LiteralPath $detailPath -NoTypeInformation -Encoding UTF8
    @($summaryRows) | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

    Publish-ReportToSharePoint `
        -FilePaths @($detailPath, $summaryPath) `
        -SiteUrl $SharePointSiteUrl `
        -FolderPath $SharePointFolderPath

    $duration = (Get-Date) - $runStart
    Write-RunbookLog 'Report complete.'
    Write-RunbookLog "  Users   : $($userRows.Count) -> $detailPath"
    Write-RunbookLog "  Summary : $($summaryRows.Count) -> $summaryPath"
    Write-RunbookLog "  SharePoint : $SharePointSiteUrl / $SharePointFolderPath"
    Write-RunbookLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)
}
catch {
    $step = if ($_.InvocationInfo.Line) { $_.InvocationInfo.Line.Trim() } else { 'unknown' }
    Write-RunbookLog "ERROR: $($_.Exception.Message)"
    Write-RunbookLog "Failed near: $step"
    if ($_.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges|AccessDenied') {
        Write-RunbookLog 'Missing Graph app permission on the Managed Identity. Ensure Reports.Read.All is assigned, then wait for token refresh.'
    }
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
