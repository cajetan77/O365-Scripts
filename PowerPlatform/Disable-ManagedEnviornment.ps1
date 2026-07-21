# 1. Sign in to your Power Platform Tenant
Add-PowerAppsAccount

# 2. Define the Target Environment ID (Get this from the Admin Center URL or Get-AdminPowerAppEnvironment)
$EnvironmentId = "Default-764b46e8-d798-4ed3-87db-ae55ed7b0432"

# 3. Create a payload mapping the protection level back to unmanaged ("Basic")
$UpdatedGovernanceConfiguration = [pscustomobject]@{ 
    protectionLevel = "Basic" 
}

# 4. Apply the payload to overwrite the active Managed Environment configuration
try {
    Write-Host "Disabling Managed Environment for ID: $EnvironmentId..." -ForegroundColor Cyan
    
    Set-AdminPowerAppEnvironmentGovernanceConfiguration `
        -EnvironmentName $EnvironmentId `
        -UpdatedGovernanceConfiguration $UpdatedGovernanceConfiguration
        
    Write-Host "Success! The environment is now unmanaged." -ForegroundColor Green
}
catch {
    Write-Warning "Failed to alter governance setup. Reason: $_"
}
