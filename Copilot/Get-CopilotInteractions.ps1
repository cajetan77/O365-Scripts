<#
.SYNOPSIS
    Export Microsoft 365 Copilot interaction data for one or all users.

.DESCRIPTION
    Two export modes are supported:

    Detailed (default)
        Uses the Graph aiInteractionHistory API. Returns prompts, responses, app class,
        accessed resources, and timestamps. Requires an Entra app with application permission
        AiEnterpriseInteraction.Read.All (+ User.Read.All to enumerate users).
        App-only (certificate) auth only — delegated permissions are not supported.

    Audit
        Uses Search-UnifiedAuditLog (Exchange Online). Returns tenant-wide interaction
        metadata (user, app, resources accessed, message IDs) but NOT prompt/response text.
        Requires Exchange Online connection and Audit Log Search permissions.

    The legacy TeamsMessagesData mailbox approach (CopilotInteractions.ps1) no longer works
    because Microsoft Graph blocks access to non-IPM folders.

.PARAMETER Mode
    Detailed = Graph aiInteractionHistory API (prompts + responses, per user).
    Audit    = Unified Audit Log (metadata only, all users in date range).

.PARAMETER UserPrincipalName
    One user to export. Omit when using -AllUsers (Detailed mode only).

.PARAMETER AllUsers
    Export every enabled member user (Detailed mode). Can be slow for large tenants.

.PARAMETER StartDate
    Start of the date range. Defaults to 30 days ago.

.PARAMETER EndDate
    End of the date range. Defaults to now.

.PARAMETER OutputPath
    CSV file path for the export. With -ExportFormat All, this is used as the base name
    (e.g. .\Report.csv produces .\Report-flat.csv, .\Report-paired.csv, .\Report-summary.csv).

.PARAMETER ExportFormat
    Flat    = One row per message/audit event (improved columns).
    Paired  = One row per prompt/response pair (Detailed mode) or per audit session (Audit mode).
    Summary = Aggregated counts by user and application.
    All     = Writes flat, paired, and summary files.

.PARAMETER IncludeRawAudit
    Audit mode only. Include the full AuditData JSON column in the flat export.

.PARAMETER ConfigPath
    Optional JSON config with TenantId, AppId, and ThumbPrint for Graph auth.

.EXAMPLE
    .\Get-CopilotInteractions.ps1 -UserPrincipalName adele@contoso.com

.EXAMPLE
    .\Get-CopilotInteractions.ps1 -Mode Audit -StartDate (Get-Date).AddDays(-7) -ExportFormat Paired

.EXAMPLE
    .\Get-CopilotInteractions.ps1 -UserPrincipalName adele@contoso.com -ExportFormat All -OutputPath .\CopilotReport.csv

.NOTES
    Graph API docs:
    https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions
    Audit log docs:
    https://learn.microsoft.com/en-us/purview/audit-copilot
#>
[CmdletBinding(DefaultParameterSetName = 'SingleUser')]
param (
    [ValidateSet('Detailed', 'Audit')]
    [string]$Mode = 'Detailed',

    [Parameter(ParameterSetName = 'SingleUser')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'AllUsers')]
    [switch]$AllUsers,

    [datetime]$StartDate = (Get-Date).AddDays(-30),
    [datetime]$EndDate = (Get-Date),

    [string]$OutputPath = ".\CopilotInteractions-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv",

    [ValidateSet('Flat', 'Paired', 'Summary', 'All')]
    [string]$ExportFormat = 'Paired',

    [switch]$IncludeRawAudit,

    [string]$ConfigPath = "D:\Powershell\O365 Scripts\SiteProvisioning\config.json",

    [string]$TenantId,
    [string]$ClientId,
    [string]$Thumbprint,

    [ValidateSet('v1.0', 'beta')]
    [string]$GraphApiVersion = 'v1.0'
)

$ErrorActionPreference = 'Stop'

function Get-ScriptGraphCredentials {
    if ($TenantId -and $ClientId -and $Thumbprint) {
        return [PSCustomObject]@{
            TenantId   = $TenantId
            ClientId   = $ClientId
            Thumbprint = $Thumbprint
        }
    }

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Config not found at '$ConfigPath'. Pass -TenantId, -ClientId, and -Thumbprint, or fix -ConfigPath."
    }

    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
    return [PSCustomObject]@{
        TenantId   = $config.TenantId
        ClientId   = $config.AppId
        Thumbprint = $config.ThumbPrint
    }
}

function Connect-GraphAppOnly {
    param($Credentials)

    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph | Out-Null
    }

    Connect-MgGraph `
        -ClientId $Credentials.ClientId `
        -TenantId $Credentials.TenantId `
        -CertificateThumbprint $Credentials.Thumbprint `
        -NoWelcome | Out-Null
}

function Get-IsoDateTime {
    param([datetime]$DateTime)
    return $DateTime.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function Get-InteractionBodyText {
    param($Interaction)

    if (-not $Interaction.body) { return '' }

    $content = [string]$Interaction.body.content
    if ([string]::IsNullOrWhiteSpace($content)) { return '' }

    if ($Interaction.body.contentType -eq 'html') {
        return ($content -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()
    }

    return $content.Trim()
}

function Get-CopilotExportPath {
    param(
        [string]$BasePath,
        [string]$Suffix
    )

    if ($BasePath -match '\.[^\\\/]+$') {
        $dir = Split-Path -Parent $BasePath
        $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
        $ext = [System.IO.Path]::GetExtension($BasePath)
        if ([string]::IsNullOrWhiteSpace($dir)) { $dir = '.' }
        return (Join-Path $dir ($name + '-' + $Suffix + $ext))
    }

    return ($BasePath + '-' + $Suffix + '.csv')
}

function Write-CopilotCsv {
    param(
        [object[]]$Data,
        [string]$Path,
        [string]$Label
    )

    $Data | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    $resolved = Resolve-Path -LiteralPath $Path
    Write-Host ("Saved {0} {1} row(s) to {2}" -f $Label, $Data.Count, $resolved)
}

function Get-CopilotAppFriendlyName {
    param(
        [string]$AppClass,
        [string]$AppIdentity,
        [string]$AppHost
    )

    if (-not [string]::IsNullOrWhiteSpace($AppHost)) { return $AppHost }

    $map = @{
        'IPM.SkypeTeams.Message.Copilot.BizChat'  = 'Microsoft 365 Chat'
        'IPM.SkypeTeams.Message.Copilot.Teams'    = 'Copilot in Teams'
        'IPM.SkypeTeams.Message.Copilot.Word'     = 'Copilot in Word'
        'IPM.SkypeTeams.Message.Copilot.Outlook'    = 'Copilot in Outlook'
        'IPM.SkypeTeams.Message.Copilot.Excel'    = 'Copilot in Excel'
        'IPM.SkypeTeams.Message.Copilot.PowerPoint' = 'Copilot in PowerPoint'
        'IPM.SkypeTeams.Message.Copilot.OneNote'    = 'Copilot in OneNote'
        'IPM.SkypeTeams.Message.Copilot.Loop'       = 'Copilot in Loop'
    }

    if ($AppClass -and $map.ContainsKey($AppClass)) {
        return $map[$AppClass]
    }

    if ($AppIdentity -match 'Copilot\.Studio') { return 'Copilot Studio' }
    if ($AppIdentity -match 'WebChat') { return 'Microsoft 365 Copilot Chat (Web)' }
    if ($AppIdentity -match 'M365Copilot') { return 'Microsoft 365 Copilot' }

    if ($AppIdentity) { return $AppIdentity }
    if ($AppClass) { return $AppClass }
    return 'Unknown'
}

function ConvertTo-CopilotPairedFromDetailed {
    param([object[]]$Rows)

    $paired = [System.Collections.Generic.List[object]]::new()

    $groups = $Rows | Group-Object -Property UserPrincipalName, RequestId
    foreach ($group in $groups) {
        $items = @($group.Group | Sort-Object CreatedDateTime)
        $prompt = $items | Where-Object { $_.InteractionType -eq 'userPrompt' } | Select-Object -First 1
        $response = $items | Where-Object { $_.InteractionType -eq 'aiResponse' } | Select-Object -First 1
        $first = if ($prompt) { $prompt } else { $items[0] }

        $paired.Add([PSCustomObject]@{
                UserPrincipalName = $first.UserPrincipalName
                UserDisplayName   = $first.UserDisplayName
                DateTime          = $first.CreatedDateTime
                Application       = (Get-CopilotAppFriendlyName -AppClass $first.AppClass -AppIdentity '' -AppHost '')
                ConversationType  = $first.ConversationType
                SessionId         = $first.SessionId
                RequestId         = $first.RequestId
                UserPrompt        = $(if ($prompt) { $prompt.Body } else { '' })
                CopilotResponse   = $(if ($response) { $response.Body } else { '' })
                Context           = $(@($items.Context | Where-Object { $_ }) | Select-Object -Unique) -join '; '
                LinkedResources   = $(@($items.LinkedResources | Where-Object { $_ }) | Select-Object -Unique) -join '; '
                LinkedResourceUrls = $(@($items.LinkedResourceUrls | Where-Object { $_ }) | Select-Object -Unique) -join '; '
            })
    }

    return $paired
}

function ConvertTo-CopilotSummary {
    param([object[]]$Rows, [string]$SourceMode)

    if ($SourceMode -eq 'Detailed') {
        return @(
            $Rows |
                Group-Object -Property UserPrincipalName, @{ Expression = { (Get-CopilotAppFriendlyName -AppClass $_.AppClass -AppIdentity '' -AppHost '') } } |
                ForEach-Object {
                    $parts = $_.Name -split ', ', 2
                    [PSCustomObject]@{
                        UserPrincipalName = $parts[0]
                        Application       = if ($parts.Count -gt 1) { $parts[1] } else { '' }
                        PromptCount       = @($_.Group | Where-Object { $_.InteractionType -eq 'userPrompt' }).Count
                        ResponseCount     = @($_.Group | Where-Object { $_.InteractionType -eq 'aiResponse' }).Count
                        ConversationCount = @($_.Group | Select-Object -ExpandProperty RequestId -Unique).Count
                        FirstActivity     = ($_.Group | Sort-Object CreatedDateTime | Select-Object -First 1).CreatedDateTime
                        LastActivity      = ($_.Group | Sort-Object CreatedDateTime -Descending | Select-Object -First 1).CreatedDateTime
                    }
                } |
                Sort-Object UserPrincipalName, Application
        )
    }

    return @(
        $Rows |
            Group-Object -Property UserPrincipalName, Application |
            ForEach-Object {
                $parts = $_.Name -split ', ', 2
                [PSCustomObject]@{
                    UserPrincipalName = $parts[0]
                    Application       = if ($parts.Count -gt 1) { $parts[1] } else { '' }
                    InteractionCount  = $_.Count
                    PromptMessageCount = ($_.Group | ForEach-Object { $_.PromptMessageCount } | Measure-Object -Sum).Sum
                    ResponseMessageCount = ($_.Group | ForEach-Object { $_.ResponseMessageCount } | Measure-Object -Sum).Sum
                    FirstActivity     = ($_.Group | Sort-Object DateTime | Select-Object -First 1).DateTime
                    LastActivity      = ($_.Group | Sort-Object DateTime -Descending | Select-Object -First 1).DateTime
                }
            } |
            Sort-Object UserPrincipalName, Application
    )
}

function Export-CopilotReportFiles {
    param(
        [object[]]$FlatRows,
        [object[]]$PairedRows,
        [object[]]$SummaryRows,
        [string]$BaseOutputPath,
        [string]$Format
    )

    $flatPath = Get-CopilotExportPath -BasePath $BaseOutputPath -Suffix 'flat'
    $pairedPath = Get-CopilotExportPath -BasePath $BaseOutputPath -Suffix 'paired'
    $summaryPath = Get-CopilotExportPath -BasePath $BaseOutputPath -Suffix 'summary'

    switch ($Format) {
        'Flat' {
            Write-CopilotCsv -Data $FlatRows -Path $flatPath -Label 'flat'
            return [PSCustomObject]@{ Flat = $flatPath }
        }
        'Paired' {
            Write-CopilotCsv -Data $PairedRows -Path $pairedPath -Label 'paired'
            return [PSCustomObject]@{ Paired = $pairedPath }
        }
        'Summary' {
            Write-CopilotCsv -Data $SummaryRows -Path $summaryPath -Label 'summary'
            return [PSCustomObject]@{ Summary = $summaryPath }
        }
        'All' {
            Write-CopilotCsv -Data $FlatRows -Path $flatPath -Label 'flat'
            Write-CopilotCsv -Data $PairedRows -Path $pairedPath -Label 'paired'
            Write-CopilotCsv -Data $SummaryRows -Path $summaryPath -Label 'summary'
            return [PSCustomObject]@{ Flat = $flatPath; Paired = $pairedPath; Summary = $summaryPath }
        }
    }
}

function ConvertTo-CopilotAuditRecord {
    param(
        $AuditResult,
        [switch]$IncludeRawAudit
    )

    $audit = $AuditResult.AuditData | ConvertFrom-Json
    $event = $audit.CopilotEventData
    if (-not $event) { $event = [PSCustomObject]@{} }

    $upn = $AuditResult.UserPrincipalName
    if ([string]::IsNullOrWhiteSpace($upn)) { $upn = $audit.UserId }

    $messages = @($event.Messages)
    $promptIds = @($messages | Where-Object { $_.isPrompt -eq $true } | ForEach-Object { $_.Id })
    $responseIds = @($messages | Where-Object { $_.isPrompt -eq $false } | ForEach-Object { $_.Id })

    $accessed = @()
    $resources = @($event.AccessedResources)
    if (-not $resources -and $audit.AccessedResources) { $resources = @($audit.AccessedResources) }
    foreach ($resource in $resources) {
        if ($resource.Name) {
            $accessed += ("{0} ({1})" -f $resource.Name, $resource.Type)
        }
        elseif ($resource.ID) {
            $accessed += $resource.ID
        }
    }

    $plugins = @($event.AISystemPlugin | ForEach-Object { $_.Name }) -join '; '
    $appHost = if ($event.AppHost) { $event.AppHost } else { $audit.AppHost }
    $appIdentity = $audit.AppIdentity

    $record = [PSCustomObject]@{
        DateTime             = $AuditResult.CreationDate
        UserPrincipalName    = $upn
        Operation            = $AuditResult.Operations
        RecordType           = $AuditResult.RecordType
        Application          = (Get-CopilotAppFriendlyName -AppClass '' -AppIdentity $appIdentity -AppHost $appHost)
        AppHost              = $appHost
        AppIdentity          = $appIdentity
        AgentName            = $audit.AgentName
        LicenseType          = $event.LicenseType
        ClientIP             = if ($audit.ClientIP) { $audit.ClientIP } else { $AuditResult.ClientIP }
        ClientRegion         = $audit.ClientRegion
        ThreadId             = $event.ThreadId
        PromptMessageIds     = ($promptIds -join '; ')
        ResponseMessageIds   = ($responseIds -join '; ')
        PromptMessageCount   = $promptIds.Count
        ResponseMessageCount = $responseIds.Count
        AccessedResources    = ($accessed -join '; ')
        PluginsUsed          = $plugins
        JailbreakDetected    = ($messages | Where-Object { $_.JailbreakDetected -eq $true }).Count -gt 0
    }

    if ($IncludeRawAudit) {
        $record | Add-Member -NotePropertyName AuditDataRaw -NotePropertyValue $AuditResult.AuditData
    }

    return $record
}

function ConvertTo-CopilotPairedFromAudit {
    param([object[]]$Rows)

    return @(
        $Rows | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName  = $_.UserPrincipalName
                DateTime           = $_.DateTime
                Application        = $_.Application
                LicenseType        = $_.LicenseType
                UserPrompt         = '[Not in audit log — use eDiscovery on mailbox or Detailed mode with Copilot license]'
                CopilotResponse    = '[Not in audit log — use eDiscovery on mailbox or Detailed mode with Copilot license]'
                PromptMessageIds   = $_.PromptMessageIds
                ResponseMessageIds = $_.ResponseMessageIds
                AccessedResources  = $_.AccessedResources
                ThreadId           = $_.ThreadId
                ClientIP           = $_.ClientIP
            }
        } | Sort-Object DateTime
    )
}

function Get-InteractionSender {
    param($Interaction)

    if ($Interaction.from.user.displayName) {
        return $Interaction.from.user.displayName
    }
    if ($Interaction.from.application.displayName) {
        return $Interaction.from.application.displayName
    }
    return ''
}

function Get-GraphErrorMessage {
    param($ErrorRecord)

    if ($ErrorRecord.ErrorDetails.Message) {
        try {
            $graphError = $ErrorRecord.ErrorDetails.Message | ConvertFrom-Json
            if ($graphError.error.message) {
                return $graphError.error.message
            }
        }
        catch {
            return $ErrorRecord.ErrorDetails.Message
        }
    }

    return $ErrorRecord.Exception.Message
}

function Invoke-GraphCopilotInteractionPage {
    param(
        [string]$RelativeUri
    )

    $response = Invoke-MgGraphRequest -Method GET -Uri $RelativeUri -OutputType PSObject
    return $response
}

function Get-GraphNextRelativeUri {
    param($Response, [string]$ApiVersion)

    $nextLink = $Response.'@odata.nextLink'
    if ([string]::IsNullOrWhiteSpace($nextLink)) { return $null }

    $prefix = "https://graph.microsoft.com/$ApiVersion"
    if ($nextLink.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        return $nextLink.Substring($prefix.Length)
    }

    return $nextLink
}

function Get-CopilotInteractionsForUser {
    param(
        [string]$UserId,
        [string]$UserDisplayName,
        [string]$Upn,
        [datetime]$RangeStart,
        [datetime]$RangeEnd,
        [string]$ApiVersion
    )

    $startIso = Get-IsoDateTime -DateTime $RangeStart
    $endIso = Get-IsoDateTime -DateTime $RangeEnd
    $filter = [uri]::EscapeDataString("createdDateTime gt $startIso and createdDateTime lt $endIso")
    $relativeUri = "/$ApiVersion/copilot/users/$UserId/interactionHistory/getAllEnterpriseInteractions?`$top=100&`$filter=$filter"

    $interactions = [System.Collections.Generic.List[object]]::new()

    do {
        $page = Invoke-GraphCopilotInteractionPage -RelativeUri $relativeUri
        foreach ($item in @($page.value)) {
            $contextNames = @($item.contexts | ForEach-Object { $_.displayName }) -join '; '
            $linkNames = @($item.links | ForEach-Object { $_.displayName }) -join '; '
            $linkUrls = @($item.links | ForEach-Object { $_.linkUrl }) -join '; '

            $interactions.Add([PSCustomObject]@{
                    UserPrincipalName = $Upn
                    UserDisplayName   = $UserDisplayName
                    CreatedDateTime   = $item.createdDateTime
                    InteractionType   = $item.interactionType
                    AppClass          = $item.appClass
                    ConversationType  = $item.conversationType
                    SessionId         = $item.sessionId
                    RequestId         = $item.requestId
                    Sender            = (Get-InteractionSender -Interaction $item)
                    Body              = (Get-InteractionBodyText -Interaction $item)
                    Context           = $contextNames
                    LinkedResources   = $linkNames
                    LinkedResourceUrls = $linkUrls
                    Locale            = $item.locale
                })
        }

        $relativeUri = Get-GraphNextRelativeUri -Response $page -ApiVersion $ApiVersion
    } while ($relativeUri)

    return $interactions
}

function Export-DetailedCopilotInteractions {
    $credentials = Get-ScriptGraphCredentials
    Connect-GraphAppOnly -Credentials $credentials

    $users = @()
    if ($AllUsers) {
        Write-Host "Loading enabled member users..."
        $users = Get-MgUser -Filter "accountEnabled eq true and userType eq 'Member'" -Property Id, DisplayName, UserPrincipalName -All
    }
    elseif ($UserPrincipalName) {
        $users = @(Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, UserPrincipalName)
    }
    else {
        throw "Detailed mode requires -UserPrincipalName or -AllUsers."
    }

    Write-Host ("Exporting detailed Copilot interactions for {0} user(s) from {1} to {2}..." -f $users.Count, $StartDate, $EndDate)

    $report = [System.Collections.Generic.List[object]]::new()
    $userIndex = 0

    foreach ($user in $users) {
        $userIndex++
        Write-Host ("[{0}/{1}] {2}" -f $userIndex, $users.Count, $user.UserPrincipalName)

        try {
            $userInteractions = Get-CopilotInteractionsForUser `
                -UserId $user.Id `
                -UserDisplayName $user.DisplayName `
                -Upn $user.UserPrincipalName `
                -RangeStart $StartDate `
                -RangeEnd $EndDate `
                -ApiVersion $GraphApiVersion

            foreach ($row in $userInteractions) {
                $report.Add($row)
            }

            Write-Host ("  Found {0} interaction record(s)." -f $userInteractions.Count)
        }
        catch {
            $reason = Get-GraphErrorMessage -ErrorRecord $_
            Write-Warning ("  Skipped {0}: {1}" -f $user.UserPrincipalName, $reason)
        }

        Start-Sleep -Milliseconds 500
    }

    if ($report.Count -eq 0) {
        Write-Warning "No interactions returned. Confirm the app has AiEnterpriseInteraction.Read.All (application) and users have Copilot licenses."
    }

    $flat = @($report | Sort-Object CreatedDateTime)
    $paired = @(ConvertTo-CopilotPairedFromDetailed -Rows $flat)
    $summary = @(ConvertTo-CopilotSummary -Rows $flat -SourceMode 'Detailed')

    Export-CopilotReportFiles -FlatRows $flat -PairedRows $paired -SummaryRows $summary -BaseOutputPath $OutputPath -Format $ExportFormat | Out-Null
    return $paired
}

function Export-AuditCopilotInteractions {
    if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false
    }

    $operations = @(
        'CopilotInteraction',
        'ConnectedAIAppInteraction',
        'AIAppInteraction'
    )

    Write-Host ("Searching Unified Audit Log from {0} to {1}..." -f $StartDate, $EndDate)

    $sessionId = [Guid]::NewGuid().ToString()
    $records = [System.Collections.Generic.List[object]]::new()
    $batch = 0

    do {
        $batch++
        $results = @(Search-UnifiedAuditLog `
            -StartDate $StartDate `
            -EndDate $EndDate `
            -Operations $operations `
            -SessionId $sessionId `
            -SessionCommand ReturnLargeSet `
            -ResultSize 5000)

        Write-Host ("  Batch {0}: {1} record(s)" -f $batch, $results.Count)

        foreach ($result in $results) {
            $records.Add((ConvertTo-CopilotAuditRecord -AuditResult $result -IncludeRawAudit:$IncludeRawAudit))
        }
    } while ($results.Count -eq 5000)

    if ($records.Count -eq 0) {
        Write-Warning "No audit records found. Confirm auditing is enabled and your account can search the audit log."
    }

    $flat = @($records | Sort-Object DateTime)
    $paired = @(ConvertTo-CopilotPairedFromAudit -Rows $flat)
    $summary = @(ConvertTo-CopilotSummary -Rows $flat -SourceMode 'Audit')

    Export-CopilotReportFiles -FlatRows $flat -PairedRows $paired -SummaryRows $summary -BaseOutputPath $OutputPath -Format $ExportFormat | Out-Null
    return $paired
}

switch ($Mode) {
    'Detailed' { Export-DetailedCopilotInteractions }
    'Audit'    { Export-AuditCopilotInteractions }
}
