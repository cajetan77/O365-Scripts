param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [datetime]$StartDate = (Get-Date).AddDays(-30),
    [datetime]$EndDate = (Get-Date),
    [string]$OutputPath,
    [int]$PageDelaySeconds = 3
)

if (-not $OutputPath) {
    $safeUser = ($UserPrincipalName -replace '[^\w\-.]', '_')
    $OutputPath = "C:\Temp\CopilotAuditDetails-$safeUser.csv"
}

$ErrorActionPreference = 'Stop'

function ConvertTo-AuditDetailRow {
    param($AuditResult)

    $audit = $AuditResult.AuditData | ConvertFrom-Json
    $event = $audit.CopilotEventData
    if (-not $event) { $event = [PSCustomObject]@{} }

    $upn = $AuditResult.UserIds
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

    $models = @($event.ModelTransparencyDetails | ForEach-Object {
        if ($_.ModelName) { "{0} ({1})" -f $_.ModelName, $_.ModelProviderName }
    }) -join '; '

    [PSCustomObject]@{
        CreationDate         = $AuditResult.CreationDate
        UserPrincipalName    = $upn
        Operation            = $AuditResult.Operations
        RecordType           = $AuditResult.RecordType
        Workload             = $audit.Workload
        AppIdentity          = $audit.AppIdentity
        AppHost              = if ($event.AppHost) { $event.AppHost } else { $audit.AppHost }
        LicenseType          = $event.LicenseType
        ClientIP             = $audit.ClientIP
        ClientRegion         = $audit.ClientRegion
        ThreadId             = $event.ThreadId
        PromptMessageCount   = $promptIds.Count
        ResponseMessageCount = $responseIds.Count
        PromptMessageIds     = ($promptIds -join '; ')
        ResponseMessageIds   = ($responseIds -join '; ')
        AccessedResources    = ($accessed -join '; ')
        PluginsUsed          = (@($event.AISystemPlugin | ForEach-Object { $_.Name }) -join '; ')
        ModelsUsed           = $models
        SensitivityLabelId   = $event.SensitivityLabelId
        JailbreakDetected    = ($messages | Where-Object { $_.JailbreakDetected -eq $true }).Count -gt 0
        MemoryUpdated        = $event.MemoryUpdated
        RecordId             = $audit.Id
        AuditDataRaw         = $AuditResult.AuditData
    }
}

if (-not (Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    Write-Error "Connect-ExchangeOnline is not installed. Please install the Exchange Online Management module."
    exit 1
}

if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowProgress $true
}

Write-Host "Searching Copilot workload audit records for $UserPrincipalName from $StartDate to $EndDate..."

$sessionId = [Guid]::NewGuid().ToString()
$allRows = [System.Collections.Generic.List[object]]::new()
$batch = 0

do {
    $batch++
    $page = @(Search-UnifiedAuditLog `
        -StartDate $StartDate `
        -EndDate $EndDate `
        -UserIds $UserPrincipalName `
        -Operations 'CopilotInteraction' `
        -SessionId $sessionId `
        -SessionCommand ReturnLargeSet `
        -ResultSize 5000)

    Write-Host ("  Page {0}: {1} record(s)" -f $batch, $page.Count)

    foreach ($record in $page) {
        $audit = $record.AuditData | ConvertFrom-Json
        if ($audit.Workload -ne 'Copilot') { continue }

        $allRows.Add((ConvertTo-AuditDetailRow -AuditResult $record))
    }

    if ($page.Count -eq 5000) {
        Start-Sleep -Seconds $PageDelaySeconds
    }
} while ($page.Count -eq 5000)

if ($allRows.Count -eq 0) {
    Write-Warning "No Copilot workload audit records returned for $UserPrincipalName in the selected date range."
}
else {
    $outputDir = Split-Path -Parent $OutputPath
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $allRows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved $($allRows.Count) record(s) to $OutputPath"
}
