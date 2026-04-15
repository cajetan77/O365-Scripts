<#


.SYNOPSIS
    Create an Entra ID App Registration and Service Principal with the necessary permissions.
.DESCRIPTION
    This script creates an Entra ID App Registration and Service Principal with the necessary permissions to access SharePoint and Microsoft Graph.
.PARAMETER DisplayName
    The display name of the app registration.
.PARAMETER CreateClientSecret
    Whether to create a client secret for the app registration.
.PARAMETER SecretDisplayName
    The display name of the client secret.
.PARAMETER SecretValidDays
    The number of days the client secret is valid.
.PARAMETER GraphAppPermissions
    The permissions to grant to the app registration for Microsoft Graph.
.PARAMETER SharePointAppPermissions
    The permissions to grant to the app registration for SharePoint.
#>



# ====== CONFIG ======
$displayName = "SharePoint Reporting"
$createClientSecret = $true
$secretDisplayName  = "runbook-secret"
$secretValidDays    = 365

#https://learn.microsoft.com/en-us/graph/permissions-reference for the permissions
$GraphAppPermissions = @(
"Sites.Read.All"
)

$SharePointAppPermissions = @(
  "Sites.Read.All" # tenant-wide SPO access
)

# ====================

# Connect to Graph with permissions to create apps/SPs
Connect-MgGraph -Scopes "Application.ReadWrite.All","Directory.ReadWrite.All" -ErrorAction Stop


$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ConsistencyLevel eventual

$sharepointSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" -ConsistencyLevel eventual


function Get-SpoAppRoleId {
    param([Parameter(Mandatory=$true)][string]$Value)
  
    $role = $sharepointSp.AppRoles | Where-Object {
      $_.Value -eq $Value -and $_.AllowedMemberTypes -contains "Application"
    }
  
    if (-not $role) { throw "SharePoint application permission not found: $Value" }
    return $role.Id
  }

function Get-GraphAppRoleId {
  param([Parameter(Mandatory=$true)][string]$Value)

  $role = $graphSp.AppRoles |
    Where-Object { $_.Value -eq $Value -and $_.AllowedMemberTypes -contains "Application" }

  if (-not $role) { throw "Graph application permission not found: $Value" }
  return $role.Id
}




# Build requiredResourceAccess (this is what "Add API permission" does)
$graphResourceAccess = @()
foreach ($perm in $GraphAppPermissions) {
  $graphResourceAccess += @{
    Id   = (Get-GraphAppRoleId -Value $perm)
    Type = "Role"   # Role = Application permission (app role)
  }
}

$spoResourceAccess = @()
foreach ($perm in $SharePointAppPermissions) {
  $spoResourceAccess += @{
    Id   = (Get-SpoAppRoleId -Value $perm)
    Type = "Role"  # application permission
  }
}

# Combine both Graph and SharePoint permissions
$requiredResourceAccess = @(
  @{
    ResourceAppId  = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
    ResourceAccess = $graphResourceAccess
  },
  @{
    ResourceAppId  = "00000003-0000-0ff1-ce00-000000000000" # SharePoint
    ResourceAccess = $spoResourceAccess
  }
)

# 1) Create or Update App Registration
$app = Get-MgApplication -Filter "DisplayName eq '$displayName'"

if ($app) {
    Write-Host "App registration already exists for $displayName - Updating permissions..." -ForegroundColor Yellow
    
    # Update the existing app registration with new permissions
    Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $requiredResourceAccess
    Write-Host "App registration permissions updated successfully" -ForegroundColor Green
} else {
    Write-Host "Creating new app registration: $displayName..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $displayName -RequiredResourceAccess $requiredResourceAccess
    Write-Host "App registration created successfully" -ForegroundColor Green
}

# 2) Get or Create the service principal (Enterprise App) so it can be used for auth/assignments
$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ConsistencyLevel eventual

if (-not $sp) {
    Write-Host "Creating service principal..." -ForegroundColor Cyan
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "Service principal created successfully" -ForegroundColor Green
} else {
    Write-Host "Service principal already exists" -ForegroundColor Yellow
}

# Helper function to check if app role assignment already exists
function Test-AppRoleAssignment {
    param(
        [Parameter(Mandatory=$true)][string]$ServicePrincipalId,
        [Parameter(Mandatory=$true)][string]$ResourceId,
        [Parameter(Mandatory=$true)][string]$AppRoleId
    )
    
    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId | 
        Where-Object { $_.ResourceId -eq $ResourceId -and $_.AppRoleId -eq $AppRoleId }
    
    return ($null -ne $existing)
}

# Grant Graph API permissions
Write-Host "Granting Graph API permissions..." -ForegroundColor Cyan
foreach ($perm in $GraphAppPermissions) {
    $appRoleId = Get-GraphAppRoleId -Value $perm
    
    if (-not (Test-AppRoleAssignment -ServicePrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $appRoleId)) {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
            principalId = $sp.Id      # your app's service principal
            resourceId  = $graphSp.Id # Microsoft Graph service principal
            appRoleId   = $appRoleId  # permission being granted
        } | Out-Null
        Write-Host "  ✓ Granted: $perm" -ForegroundColor Green
    } else {
        Write-Host "  ⊙ Already granted: $perm" -ForegroundColor Gray
    }
}

# Grant SharePoint API permissions
Write-Host "Granting SharePoint API permissions..." -ForegroundColor Cyan
foreach ($perm in $SharePointAppPermissions) {
    $appRoleId = Get-SpoAppRoleId -Value $perm
    
    if (-not (Test-AppRoleAssignment -ServicePrincipalId $sp.Id -ResourceId $sharepointSp.Id -AppRoleId $appRoleId)) {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
            principalId = $sp.Id           # your app's service principal
            resourceId  = $sharepointSp.Id # SharePoint service principal
            appRoleId   = $appRoleId       # permission being granted
        } | Out-Null
        Write-Host "  ✓ Granted: $perm" -ForegroundColor Green
    } else {
        Write-Host "  ⊙ Already granted: $perm" -ForegroundColor Gray
    }
}

# 3) (Optional) Create a client secret
$secretValue = $null
if ($createClientSecret) {
    # Check if secret with same name already exists
    $existingSecrets = Get-MgApplication -ApplicationId $app.Id | Select-Object -ExpandProperty PasswordCredentials
    
    if (-not ($existingSecrets | Where-Object { $_.DisplayName -eq $secretDisplayName })) {
        Write-Host "Creating new client secret..." -ForegroundColor Cyan
        $endDate = (Get-Date).AddDays($secretValidDays)

        $pwdCred = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            displayName = $secretDisplayName
            endDateTime = $endDate
        }

        # IMPORTANT: This is the only time you can read the secret value
        $secretValue = $pwdCred.SecretText
        Write-Host "Client secret created successfully" -ForegroundColor Green
    } else {
        Write-Host "Client secret '$secretDisplayName' already exists. Skipping creation." -ForegroundColor Yellow
        Write-Host "Note: Existing secret values cannot be retrieved. Create a new secret with a different name if needed." -ForegroundColor Yellow
    }
}

# Output results
$result = [PSCustomObject]@{
    DisplayName        = $app.DisplayName
    ClientId           = $app.AppId
    AppObjectId        = $app.Id
    ServicePrincipalId = $sp.Id
    TenantId           = (Get-MgContext).TenantId
    ClientSecret       = $secretValue
}

$result | Format-List












