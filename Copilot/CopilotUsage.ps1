#Install-Module ExchangeOnlineManagement -Scope CurrentUser

<#
.SYNOPSIS
    Summarize Copilot interaction usage from the Unified Audit Log.

.NOTES
    Search-UnifiedAuditLog is rate-limited. This script uses session pagination,
    retries with backoff, and pacing between requests to avoid 429 errors.
#>
param(
    [datetime]$StartDate = (Get-Date).AddDays(-30).Date,
    [datetime]$EndDate = (Get-Date),
    [string]$RawOutput = "C:\Temp\CopilotChatAuditRaw.csv",
    [string]$SummaryOutput = "C:\Temp\CopilotChatUpgradeCandidates.csv",
    [int]$ChunkDays = 7,
    [int]$PageDelaySeconds = 3,
    [int]$ChunkDelaySeconds = 60,
    [int]$MaxRetries = 6
)

$ErrorActionPreference = 'Stop'

function Get-CopilotAppHostFromAuditData {
    param([string]$AuditDataJson)

    $data = $AuditDataJson | ConvertFrom-Json
    if ($data.CopilotEventData.AppHost) { return [string]$data.CopilotEventData.AppHost }
    if ($data.AppHost) { return [string]$data.AppHost }
    return ''
}

function Invoke-UnifiedAuditLogSearch {
    param(
        [datetime]$WindowStart,
        [datetime]$WindowEnd,
        [string]$SessionId,
        [int]$MaxRetries
    )

    $attempt = 0
    while ($true) {
        $attempt++
        $warnings = @()

        try {
            $results = @(Search-UnifiedAuditLog `
                -StartDate $WindowStart `
                -EndDate $WindowEnd `
                -Operations 'CopilotInteraction' `
                -SessionId $SessionId `
                -SessionCommand ReturnLargeSet `
                -ResultSize 5000 `
                -WarningVariable warnings `
                -WarningAction SilentlyContinue)
        }
        catch {
            $message = $_.Exception.Message
            if ($message -notmatch 'TooManyRequests|429|Too many requests|throttl') {
                throw
            }
            $results = @()
        }

        $throttled = @($warnings | Where-Object {
            $_ -match 'TooManyRequests|429|Too many requests|throttl'
        }).Count -gt 0

        if (-not $throttled -or $results.Count -gt 0) {
            return $results
        }

        if ($attempt -ge $MaxRetries) {
            throw "Audit search throttled after $MaxRetries attempts for window $WindowStart to $WindowEnd. Wait 10-15 minutes and rerun, or increase -ChunkDelaySeconds."
        }

        $delay = [math]::Min(120, [math]::Pow(2, $attempt) + (Get-Random -Minimum 1 -Maximum 5))
        Write-Warning "Throttled (attempt $attempt/$MaxRetries). Waiting ${delay}s before retry..."
        Start-Sleep -Seconds $delay
    }
}

function Get-CopilotAuditRecords {
    param(
        [datetime]$RangeStart,
        [datetime]$RangeEnd,
        [int]$PageDelaySeconds,
        [int]$MaxRetries
    )

    $sessionId = [Guid]::NewGuid().ToString()
    $records = [System.Collections.Generic.List[object]]::new()
    $batch = 0

    do {
        $batch++
        $page = Invoke-UnifiedAuditLogSearch `
            -WindowStart $RangeStart `
            -WindowEnd $RangeEnd `
            -SessionId $sessionId `
            -MaxRetries $MaxRetries

        Write-Host ("  Page {0}: {1} record(s)" -f $batch, $page.Count)

        foreach ($r in $page) {
            $records.Add([PSCustomObject]@{
                UserUPN      = $r.UserIds
                CreationDate = $r.CreationDate
                AppHost      = (Get-CopilotAppHostFromAuditData -AuditDataJson $r.AuditData)
                Workload     = $r.Workload
                Operation    = $r.Operations
            })
        }

        if ($page.Count -eq 5000) {
            Start-Sleep -Seconds $PageDelaySeconds
        }
    } while ($page.Count -eq 5000)

    if ($records.Count -ge 50000) {
        Write-Warning "Window returned 50,000+ records (audit search cap). Narrow the date range or reduce -ChunkDays."
    }

    return $records
}

if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false
}

$allRecords = [System.Collections.Generic.List[object]]::new()
$current = $StartDate

Write-Host "Collecting CopilotInteraction audit records from $StartDate to $EndDate..."

while ($current -lt $EndDate) {
    $windowStart = $current
    $windowEnd = $current.AddDays($ChunkDays)
    if ($windowEnd -gt $EndDate) { $windowEnd = $EndDate }

    Write-Host "Searching $windowStart to $windowEnd"

    $chunkRecords = Get-CopilotAuditRecords `
        -RangeStart $windowStart `
        -RangeEnd $windowEnd `
        -PageDelaySeconds $PageDelaySeconds `
        -MaxRetries $MaxRetries

    foreach ($record in $chunkRecords) {
        $allRecords.Add($record)
    }

    Write-Host ("  Chunk total so far: {0}" -f $allRecords.Count)

    $current = $windowEnd
    if ($current -lt $EndDate) {
        Start-Sleep -Seconds $ChunkDelaySeconds
    }
}

if ($allRecords.Count -eq 0) {
    Write-Warning "No records returned. If you saw TooManyRequests errors, wait 10-15 minutes and rerun."
}

$allRecords | Export-Csv $RawOutput -NoTypeInformation
Write-Host "Saved $($allRecords.Count) raw record(s) to $RawOutput"

$summary = $allRecords |
    Group-Object UserUPN |
    ForEach-Object {
        $events = $_.Group
        $activeDays = ($events.CreationDate | ForEach-Object { ([datetime]$_).Date } | Select-Object -Unique).Count

        [PSCustomObject]@{
            UserUPN          = $_.Name
            InteractionCount = $_.Count
            ActiveDays       = $activeDays
            LastActivity     = ($events.CreationDate | Sort-Object -Descending | Select-Object -First 1)
            AppsUsed         = ($events.AppHost | Where-Object { $_ } | Select-Object -Unique) -join ', '
            Recommendation   = if ($_.Count -ge 250 -and $activeDays -ge 15) {
                'Upgrade Now'
            }
            elseif ($_.Count -ge 100 -and $activeDays -ge 10) {
                'Review'
            }
            else {
                'Monitor'
            }
        }
    }

$summary |
    Sort-Object InteractionCount -Descending |
    Export-Csv $SummaryOutput -NoTypeInformation

Write-Host "Saved summary for $($summary.Count) user(s) to $SummaryOutput"
