# --- CONFIGURATION ---

$OutputDirectory = Get-Location
$ConfigPath = Join-Path $OutputDirectory 'config.json'
$ListName = 'PowerApps'
$ReportPath = "$($OutputDirectory.Path)\PowerApps_Inventory_Report.csv"

function ConvertTo-SharePointText {
    param([string]$Value)

    if ([string]::IsNullOrEmpty($Value)) { return $Value }

    $cleaned = -join ($Value.ToCharArray() | ForEach-Object {
            $code = [int][char]$_
            if ($code -eq 0x9 -or $code -eq 0xA -or $code -eq 0xD -or ($code -ge 0x20 -and $code -le 0xD7FF) -or ($code -ge 0xE000 -and $code -le 0xFFFD)) {
                $_
            }
            elseif ($code -lt 0x20) {
                ' '
            }
        })

    return ($cleaned -replace '\s+', ' ').Trim()
}

function ConvertTo-SharePointDateTime {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $formats = @('dd-MM-yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm:ss', 'MM/dd/yyyy HH:mm:ss')
    foreach ($format in $formats) {
        try {
            return [datetime]::ParseExact($Value, $format, [cultureinfo]::InvariantCulture)
        }
        catch {
            # try next format
        }
    }

    [datetime]$parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [ref]$parsed)) { return $parsed }
    return $null
}

# --- AUTHENTICATION & INITIALIZATION ---
Write-Host "Authenticating to Power Platform Admin..." -ForegroundColor Cyan
Add-PowerAppsAccount

Write-Host "Fetching all environments..." -ForegroundColor Cyan
$Environments = Get-AdminPowerAppEnvironment

$UnusedAppsReport = @()

# --- PROCESSING ---
$envCount = 0
foreach ($Env in $Environments) {
    $envCount++
    Write-Host "Scanning environment $envCount of $($Environments.Count): $($Env.DisplayName)" -ForegroundColor Green
    $EnvId = $Env.EnvironmentName
    $EnvName = $Env.Internal.properties.displayName
    Write-Host "Scanning environment: $EnvName..." -ForegroundColor Green

    try {
        $Apps = Get-AdminPowerApp -EnvironmentName $EnvId -ErrorAction Stop
    }
    catch {
        Write-Host "Skipping $EnvName due to access restrictions or errors." -ForegroundColor Yellow
        continue
    }
    $appCount = 0
    foreach ($App in $Apps) {
        $appCount++
        Write-Host "Scanning App $appCount of $($Apps.Count): $($App.DisplayName) in environment: $EnvName" -ForegroundColor Green
        $LastModified = [DateTime]$App.LastModifiedTime

        # if ($LastModified -ge $CutoffDate) { continue }

        $coOwners = @()        
        $sharedWith = @()
        try {
            $permissions = Get-AdminPowerAppRoleAssignment -EnvironmentName $EnvId -AppName $App.AppName -ErrorAction Stop
            $coOwners = @($permissions | Where-Object { $_.RoleType -eq 'CanEdit' } | ForEach-Object {
                    if ($_.PrincipalDisplayName) { $_.PrincipalDisplayName } else { $_.PrincipalObjectId }
                })
            $sharedWith = @($permissions | Where-Object { $_.RoleType -in @('CanView', 'CanViewWithShare') } | ForEach-Object {
                    if ($_.PrincipalDisplayName) { $_.PrincipalDisplayName } else { $_.PrincipalObjectId }
                })
        }
        catch {
            Write-Host "Could not retrieve permissions for $($App.DisplayName): $_" -ForegroundColor Yellow
        }

        $UnusedAppsReport += [PSCustomObject]@{
            EnvironmentName  = ConvertTo-SharePointText $EnvName
            EnvironmentID    = $EnvId
            AppDisplayName   = ConvertTo-SharePointText $App.DisplayName
            Owner            = ConvertTo-SharePointText $App.Owner.displayName
            OwnerEmail       = $App.Owner.userPrincipalName
            CoOwner          = ConvertTo-SharePointText ($coOwners -join ' | ')
            SharedWith       = ConvertTo-SharePointText ($sharedWith -join ' | ')
            LastModifiedTime = $LastModified
            AppID            = $App.AppName
            AppType          = ConvertTo-SharePointText $App.Internal.appType
        }
    }
}

# --- OUTPUT GENERATION ---


$UnusedAppsReport | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "Scan complete! Found $($UnusedAppsReport.Count) unused apps." -ForegroundColor Green
Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Warning "Config not found at $ConfigPath. Skipping SharePoint upload."
    return
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
Connect-PnPOnline -Url $config.SiteUrl -ClientId $config.AppId -Tenant $config.TenantId -Thumbprint $config.ThumbPrint

$csv = Import-Csv -Path $ReportPath
$existingItemsByAppId = @{}
Get-PnPListItem -List $ListName -Fields 'AppID' -PageSize 5000 | ForEach-Object {
    $appId = $_.FieldValues.AppID
    if (-not [string]::IsNullOrWhiteSpace($appId)) {
        $existingItemsByAppId[$appId] = $_.Id
    }
}

$added = 0
$updated = 0
foreach ($item in $csv) {
    $values = @{
        Title           = ConvertTo-SharePointText $item.AppDisplayName
        EnvironmentName = ConvertTo-SharePointText $item.EnvironmentName
        EnvironmentID   = $item.EnvironmentID
        AppDisplayName  = ConvertTo-SharePointText $item.AppDisplayName
        CoOwner         = ConvertTo-SharePointText $item.CoOwner
        SharedWith      = ConvertTo-SharePointText $item.SharedWith
        AppID           = $item.AppID
        AppType         = ConvertTo-SharePointText $item.AppType
    }

    $lastModified = ConvertTo-SharePointDateTime -Value $item.LastModifiedTime
    if ($lastModified) { $values['LastModifiedTime'] = $lastModified }

    if (-not [string]::IsNullOrWhiteSpace($item.OwnerEmail)) {
        $values['Owner'] = $item.OwnerEmail
    }

    try {
        if ($existingItemsByAppId.ContainsKey($item.AppID)) {
            Set-PnPListItem -List $ListName -Identity $existingItemsByAppId[$item.AppID] -Values $values -ErrorAction Stop
            $updated++
        }
        else {
            $newItem = Add-PnPListItem -List $ListName -Values $values -ErrorAction Stop
            $existingItemsByAppId[$item.AppID] = $newItem.Id
            $added++
        }
    }
    catch {
        Write-Warning "Failed to sync '$($item.AppDisplayName)' to '$ListName': $_"
    }
}

Write-Host "Synced $($csv.Count) apps to SharePoint list '$ListName' ($added added, $updated updated)." -ForegroundColor Green
