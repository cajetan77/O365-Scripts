# SignIns requires Microsoft.Graph.Authentication at the same version (see .psd1 RequiredModules).
# Only one Authentication assembly can load per session — import Graph modules before Connect-MgGraph.
$graphVersion = '2.37.0'

function Initialize-GraphModules {
    param([string]$Version)

    $loadedAuth = Get-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    if ($loadedAuth -and $loadedAuth.Version.ToString() -ne $Version) {
        Write-Warning "Microsoft.Graph.Authentication $($loadedAuth.Version) is already loaded. Unloading Graph modules for $Version."
        if (Get-MgContext -ErrorAction SilentlyContinue) { Disconnect-MgGraph | Out-Null }
        Get-Module Microsoft.Graph* -ErrorAction SilentlyContinue | Remove-Module -Force -ErrorAction SilentlyContinue
    }

    foreach ($root in ($env:PSModulePath -split ';' | Where-Object { $_ -and (Test-Path $_) })) {
        $authRoot = Join-Path $root 'Microsoft.Graph.Authentication'
        if (-not (Test-Path $authRoot)) { continue }
        Get-ChildItem $authRoot -Directory | Where-Object Name -ne $Version | ForEach-Object {
            Write-Warning "Removing stale Microsoft.Graph.Authentication $($_.Name) from $($_.FullName)"
            Remove-Item $_.FullName -Recurse -Force
        }
    }

    if (-not (Get-Module Microsoft.Graph.Authentication -ListAvailable | Where-Object Version -eq $Version)) {
        Install-Module Microsoft.Graph.Authentication -RequiredVersion $Version -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Authentication -RequiredVersion $Version -Force
    Import-Module Microsoft.Graph.Identity.SignIns -RequiredVersion $Version -Force
}

Initialize-GraphModules -Version $graphVersion

$configPath = "D:\Powershell\O365 Scripts\SiteProvisioning\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId
$TenantName = $config.TenantName

$Thumbprint = $config.ThumbPrint
$clientSecret = $config.ClientSecret

Connect-MgGraph -ClientId $ClientId -TenantId $tenantId -CertificateThumbprint $Thumbprint -NoWelcome
$context = Get-MgContext

Get-MgPolicyAuthorizationPolicy | Select-Object allowedToSignUpEmailBasedSubscriptions, allowEmailVerifiedUsersToJoinOrganization

# To block self-service signup (requires Policy.ReadWrite.Authorization on the app):
# Update-MgPolicyAuthorizationPolicy -AllowedToSignUpEmailBasedSubscriptions:$false

#$policies = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where-Object { $_.ProductName -eq "Microsoft 365 Copilot" }

#foreach ($policy in $policies) {
#   Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $policy.ProductId -Value "Disabled"
#}



$param = @{
    allowedToSignUpEmailBasedSubscriptions    = $false
    allowEmailVerifiedUsersToJoinOrganization = $false
}
Update-MgPolicyAuthorizationPolicy -BodyParameter $param


Disconnect-MgGraph
