 
param(
    [Parameter(Mandatory = $false)]
    [string]$DisplayName,
    [Parameter(Mandatory = $false)]
    [string]$Alias,
    [Parameter(Mandatory = $false)]
    [string]$OwnerUpn,
    [Parameter(Mandatory = $false)]
    [string]$OwnerDisplayName
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
 
Connect-MgGraph -ClientId $ClientId -TenantId $tenantId -CertificateThumbprint $Thumbprint  -NoWelcome
$context = Get-MgContext
 
$user = Get-MgUser -Filter "DisplayName eq '$OwnerDisplayName'"
$userId = $user.Id

 

$existinggroup = Get-MgGroup -Filter " mailnickname eq '$ShortTitle'"
if ($existinggroup) {
    Write-Warning "A group with the alias '$ShortTitle' already exists. Exiting script."
    
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
 
 
 