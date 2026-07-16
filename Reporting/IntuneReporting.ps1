<#
.SYNOPSIS
    Exports all Entra ID devices using Get-MgDevice, including assigned users.

.DESCRIPTION
    Read-only inventory of directory devices via Microsoft Graph v1.0.
    Uses device list with registeredOwners / registeredUsers expanded so each
    device shows who it is assigned to.

.NOTES
    Application permissions (admin consent):
      Device.Read.All
      User.Read.All   (recommended so owner UPN/displayName are returned)
#>
[CmdletBinding()]
Param(
    [switch]$EnabledDevicesOnly,
    [int]$StaleDays = 30,
    [int]$GraphRetries = 5,
    [int]$GraphRetryDelaySeconds = 10,

    [string]$TenantId = '764b46e8-d798-4ed3-87db-ae55ed7b0432',
    [string]$ClientId = 'dc223b11-5ab5-4a33-988a-3474b25eb9be',
    [string]$ClientSecret = '',
    [string]$Thumbprint = '800DB610ED947E9251A199ABFEA40AED1738128E',
    [string]$ExportPath = '.\Intune-Report'
)

$ErrorActionPreference = 'Stop'
$runStart = Get-Date

function Write-ReportLog {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function Connect-MgGraphApp {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    Write-ReportLog 'Connecting to Microsoft Graph...'
    if ($Thumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    }
    elseif ($ClientSecret) {
        $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential ([PSCredential]::new($ClientId, $secureSecret)) -NoWelcome
    }
    else {
        throw 'Provide -Thumbprint or -ClientSecret.'
    }

    Set-MgRequestContext -Retries $GraphRetries -RetryDelay $GraphRetryDelaySeconds | Out-Null
}

function Get-DirectoryUserLabel {
    param($Person)

    if (-not $Person) { return $null }

    $upn = $Person.userPrincipalName
    $name = $Person.displayName
    $id = $Person.id

    if ($upn -and $name) { return "$name <$upn>" }
    if ($upn) { return $upn }
    if ($name) { return $name }
    if ($id) { return $id }
    return $null
}

function Get-AssignmentInfo {
    param($Device)

    $owners = @($Device.registeredOwners)
    $users = @($Device.registeredUsers)

    $ownerLabels = @(
        $owners | ForEach-Object { Get-DirectoryUserLabel -Person $_ } | Where-Object { $_ }
    )
    $userLabels = @(
        $users | ForEach-Object { Get-DirectoryUserLabel -Person $_ } | Where-Object { $_ }
    )

    $primary = if ($ownerLabels.Count -gt 0) {
        $ownerLabels[0]
    }
    elseif ($userLabels.Count -gt 0) {
        $userLabels[0]
    }
    else {
        ''
    }

    return [PSCustomObject]@{
        AssignedTo             = $primary
        RegisteredOwners       = ($ownerLabels -join '; ')
        RegisteredOwnerUPNs    = (($owners | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ }) -join '; ')
        RegisteredUsers        = ($userLabels -join '; ')
        RegisteredUserUPNs     = (($users | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ }) -join '; ')
        HasAssignment          = [bool]($primary)
    }
}

function Get-AllDevicesWithAssignment {
    param([int]$PageSize = 100)

    # Keep page size modest when expanding owners/users (Graph payload size).
    $devices = [System.Collections.Generic.List[object]]::new()
    $select = 'id,deviceId,displayName,accountEnabled,operatingSystem,operatingSystemVersion,trustType,isManaged,isCompliant,profileType,deviceOwnership,enrollmentType,managementType,manufacturer,model,approximateLastSignInDateTime,registrationDateTime,createdDateTime,deviceCategory'
    $expand = 'registeredOwners($select=id,displayName,userPrincipalName),registeredUsers($select=id,displayName,userPrincipalName)'
    $uri = "https://graph.microsoft.com/v1.0/devices?`$top=$PageSize&`$select=$select&`$expand=$expand"

    do {
        $page = Invoke-MgGraphRequest -Uri $uri -Method GET
        if ($page.value) {
            $devices.AddRange(@($page.value))
        }

        Write-Progress -Activity 'Retrieving devices (Get-MgDevice /devices)' `
            -Status "$($devices.Count) device(s) retrieved" `
            -PercentComplete $(if ($page.'@odata.nextLink') { -1 } else { 100 })

        $uri = $page.'@odata.nextLink'
    } while ($uri)

    Write-Progress -Activity 'Retrieving devices (Get-MgDevice /devices)' -Completed
    return @($devices)
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $Data | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

try {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $exportRoot = [System.IO.Path]::GetFullPath($ExportPath)
    if (-not (Test-Path -LiteralPath $exportRoot)) {
        New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
    }

    $devicesPath = Join-Path $exportRoot "Devices-$timestamp.csv"
    $summaryPath = Join-Path $exportRoot "DeviceSummary-$timestamp.csv"
    $runLogPath = Join-Path $exportRoot "RunLog-$timestamp.csv"
    $script:LogPath = Join-Path $exportRoot "RunLog-$timestamp.log"

    Write-ReportLog "Starting device report (TenantId: $TenantId)"
    Connect-MgGraphApp

    Write-ReportLog 'Retrieving devices with owners/users (Graph /devices + expand)...'
    $allDevices = Get-AllDevicesWithAssignment
    Write-ReportLog "Retrieved $($allDevices.Count) device(s)."

    if ($EnabledDevicesOnly) {
        $allDevices = @($allDevices | Where-Object { $_.accountEnabled -eq $true })
        Write-ReportLog "Enabled devices: $($allDevices.Count)"
    }

    $deviceRows = [System.Collections.Generic.List[object]]::new()
    $osCounts = @{}
    $trustCounts = @{}
    $staleCount = 0
    $disabledCount = 0
    $unassignedCount = 0
    $processed = 0

    foreach ($d in $allDevices) {
        $processed++
        if ($processed % 1000 -eq 0 -or $processed -eq $allDevices.Count) {
            Write-Progress -Activity 'Building report' -Status "$processed / $($allDevices.Count)" `
                -PercentComplete (($processed / [Math]::Max($allDevices.Count, 1)) * 100)
        }

        $assignment = Get-AssignmentInfo -Device $d

        $lastSignIn = $null
        if ($d.approximateLastSignInDateTime) {
            $lastSignIn = [datetime]$d.approximateLastSignInDateTime
        }

        $daysSinceSignIn = if ($lastSignIn) {
            [int](New-TimeSpan -Start $lastSignIn -End (Get-Date)).TotalDays
        }
        else { $null }

        $os = if ($d.operatingSystem) { [string]$d.operatingSystem } else { 'Unknown' }
        $trust = if ($d.trustType) { [string]$d.trustType } else { 'unknown' }

        if ($d.accountEnabled -eq $false) { $disabledCount++ }
        if (-not $assignment.HasAssignment) { $unassignedCount++ }
        if ($daysSinceSignIn -ne $null -and $daysSinceSignIn -gt $StaleDays) { $staleCount++ }

        if (-not $osCounts.ContainsKey($os)) { $osCounts[$os] = 0 }
        if (-not $trustCounts.ContainsKey($trust)) { $trustCounts[$trust] = 0 }
        $osCounts[$os]++
        $trustCounts[$trust]++

        $deviceRows.Add([PSCustomObject]@{
            DisplayName                   = $d.displayName
            AssignedTo                    = $assignment.AssignedTo
            RegisteredOwners              = $assignment.RegisteredOwners
            RegisteredOwnerUPNs           = $assignment.RegisteredOwnerUPNs
            RegisteredUsers               = $assignment.RegisteredUsers
            RegisteredUserUPNs            = $assignment.RegisteredUserUPNs
            DeviceId                      = $d.deviceId
            ObjectId                      = $d.id
            AccountEnabled                = $d.accountEnabled
            OperatingSystem               = $d.operatingSystem
            OperatingSystemVersion        = $d.operatingSystemVersion
            TrustType                     = $d.trustType
            IsManaged                     = $d.isManaged
            IsCompliant                   = $d.isCompliant
            ProfileType                   = $d.profileType
            DeviceOwnership               = $d.deviceOwnership
            EnrollmentType                = $d.enrollmentType
            ManagementType                = $d.managementType
            Manufacturer                  = $d.manufacturer
            Model                         = $d.model
            ApproximateLastSignInDateTime = $d.approximateLastSignInDateTime
            DaysSinceLastSignIn           = $daysSinceSignIn
            RegistrationDateTime          = $d.registrationDateTime
            CreatedDateTime               = $d.createdDateTime
            DeviceCategory                = $d.deviceCategory
        })
    }

    Write-Progress -Activity 'Building report' -Completed

    $summaryRows = [System.Collections.Generic.List[object]]::new()
    $summaryRows.Add([PSCustomObject]@{ Metric = 'ReportGeneratedUtc'; Value = (Get-Date).ToUniversalTime().ToString('o') })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TenantId'; Value = $TenantId })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'Source'; Value = 'Get-MgDevice (/devices with registeredOwners/Users)' })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'TotalDevices'; Value = $deviceRows.Count })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'AssignedDevices'; Value = ($deviceRows.Count - $unassignedCount) })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'UnassignedDevices'; Value = $unassignedCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = 'DisabledDevices'; Value = $disabledCount })
    $summaryRows.Add([PSCustomObject]@{ Metric = "StaleOver${StaleDays}Days"; Value = $staleCount })

    foreach ($entry in ($osCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "OS:$($entry.Key)"; Value = $entry.Value })
    }
    foreach ($entry in ($trustCounts.GetEnumerator() | Sort-Object Name)) {
        $summaryRows.Add([PSCustomObject]@{ Metric = "TrustType:$($entry.Key)"; Value = $entry.Value })
    }

    Export-ReportCsv -Data $deviceRows -Path $devicesPath
    Export-ReportCsv -Data $summaryRows -Path $summaryPath

    $duration = (Get-Date) - $runStart
    $runMeta = [PSCustomObject]@{
        ReportGeneratedUtc = (Get-Date).ToUniversalTime().ToString('o')
        TenantId           = $TenantId
        TotalDevices       = $deviceRows.Count
        AssignedDevices    = ($deviceRows.Count - $unassignedCount)
        UnassignedDevices  = $unassignedCount
        DurationMinutes    = [Math]::Round($duration.TotalMinutes, 2)
        Status             = 'Success'
        DevicesCsv         = $devicesPath
        SummaryCsv         = $summaryPath
    }
    Export-ReportCsv -Data $runMeta -Path $runLogPath

    Write-ReportLog 'Report complete.'
    Write-ReportLog "  Devices : $($deviceRows.Count) -> $devicesPath"
    Write-ReportLog "  Summary : $($summaryRows.Count) -> $summaryPath"
    Write-ReportLog ("Duration: {0:N1} minute(s)" -f $duration.TotalMinutes)

    if ($deviceRows.Count -eq 0) {
        Write-ReportLog 'WARNING: No devices returned. Confirm Device.Read.All (+ User.Read.All) with admin consent.'
    }
}
catch {
    Write-ReportLog "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
