 
param(
    [Parameter(Mandatory = $true)]
    [string]$DisplayName,
    [Parameter(Mandatory = $true)]
    [string]$Alias,
    [Parameter(Mandatory = $true)]
    [string]$OwnerUpn,
    [Parameter(Mandatory = $true)]
    [string]$OwnerDisplayName,
    [Parameter(Mandatory = $true)]
    [string]$SvcAccountName
)
# Load configuration from JSON file
# Load configuration from JSON file
$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$TenantName = $config.TenantName    
$Thumbprint = $config.ThumbPrint
$clientSecret = $config.ClientSecret

# SharePoint URLs
$AdminUrl = "https://$tenantname-admin.sharepoint.com"   # e.g. https://contoso-admin.sharepoint.com

   
# ===========================
# BEGIN SCRIPT
# ===========================
$Title = $DisplayName.Trim()
$ShortTitle = $Alias.Trim()
 
Import-Module Microsoft.Graph.Authentication 
 
try {
    Connect-MgGraph -ClientId $ClientId -TenantId $tenantId -CertificateThumbprint $Thumbprint  -NoWelcome
    $context = Get-MgContext
}
catch {
    Write-Host "Error connecting to Microsoft Graph: $_"
    exit
}


try {
    $user = Get-MgUser -Filter "DisplayName eq '$OwnerDisplayName'"
    $userId = $user.Id
}
catch {
    Write-Host "Error retrieving user information from Microsoft Graph: $_"
    
    $user = Get-MgUser -Filter "DisplayName eq '$SvcAccountName'"
    $userId = $user.Id
    
}


 

$existinggroup = Get-MgGroup -Filter " mailnickname eq '$ShortTitle'"
if ($existinggroup) {
    Write-Warning "A group with the alias '$ShortTitle' already exists. Exiting script."
    exit
}


Write-Host "Owner Email $($user.UserPrincipalName)"
Write-Host "Owner Id $($userId)"

try {
    $group = New-MgGroup -DisplayName $Title `
        -MailNickname $ShortTitle `
        -Description "$Title Group" `
        -GroupTypes @("Unified") `
        -MailEnabled: $true -SecurityEnabled:$false -Visibility "Private" 
    Start-Sleep 30s
    $g = Get-MgGroup -GroupId $group.Id
  
}
catch {
    Write-Host "Error creating the Office 365 Group: $_"
    exit
}
 
 
 
try {
    $newGroupOwner = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/{$userId}"
    }
   
    New-MgGroupOwnerByRef -GroupId $groupId -BodyParameter $newGroupOwner
    Write-Host "Added $($user.UserPrincipalName) as owner to the group."
}
catch {
    Write-Host "Error adding owner to the Office 365 Group: $_"
}
 
 
Start-Sleep 30s
              
try {
    $params = @{
        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        "group@odata.bind"    = "https://graph.microsoft.com/v1.0/groups('$groupId')"
    }
   
    Update-MgGroup -GroupId $groupId -BodyParameter @{
        hideFromAddressLists   = $true
        hideFromOutlookClients = $true
    }
   
    New-MgTeam -BodyParameter $params

    Write-Host "Team created successfully for the Office 365 Group."
   
}
catch {
    Write-Host "Error creating the Team for the Office 365 Group: $_"
}
 
 
 