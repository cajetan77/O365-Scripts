<#
.SYNOPSIS
    Setup script for Azure Automation App Expiry Monitor
.DESCRIPTION
    This script helps set up the Azure Automation environment for the App Expiry Monitor runbook.
    It provides guidance and PowerShell commands for configuring the necessary components.
#>

Write-Host "=== Azure Automation Setup Guide for App Expiry Monitor ===" -ForegroundColor Cyan
Write-Host ""

Write-Host "STEP 1: Create Azure Automation Account" -ForegroundColor Yellow
Write-Host "---------------------------------------"
Write-Host "1. Go to Azure Portal > Create a resource > Automation Account"
Write-Host "2. Choose your subscription, resource group, and region"
Write-Host "3. Enable System-assigned managed identity"
Write-Host "4. Create the Automation Account"
Write-Host ""

Write-Host "STEP 2: Configure Managed Identity Permissions" -ForegroundColor Yellow
Write-Host "----------------------------------------------"
Write-Host "The Managed Identity needs the following Microsoft Graph permissions:"
Write-Host "• Application.Read.All (to read app registrations)"
Write-Host "• Sites.ReadWrite.All (to update SharePoint lists)"
Write-Host ""
Write-Host "PowerShell commands to grant permissions (run as Global Admin):"
Write-Host ""
Write-Host @"
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

# Get your Automation Account's Managed Identity
`$automationAccountName = "YourAutomationAccountName"
`$resourceGroupName = "YourResourceGroupName"
`$subscriptionId = "YourSubscriptionId"

# Get the Managed Identity Object ID from Azure Portal > Automation Account > Identity
`$managedIdentityObjectId = "PASTE-OBJECT-ID-HERE"

# Get Microsoft Graph Service Principal
`$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# Get required app roles
`$applicationReadAllRole = `$graphServicePrincipal.AppRoles | Where-Object {`$_.Value -eq "Application.Read.All"}
`$sitesReadWriteAllRole = `$graphServicePrincipal.AppRoles | Where-Object {`$_.Value -eq "Sites.ReadWrite.All"}

# Grant permissions
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId `$managedIdentityObjectId -PrincipalId `$managedIdentityObjectId -ResourceId `$graphServicePrincipal.Id -AppRoleId `$applicationReadAllRole.Id
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId `$managedIdentityObjectId -PrincipalId `$managedIdentityObjectId -ResourceId `$graphServicePrincipal.Id -AppRoleId `$sitesReadWriteAllRole.Id

Write-Host "Permissions granted successfully!" -ForegroundColor Green
"@
Write-Host ""

Write-Host "STEP 3: Import Required PowerShell Modules" -ForegroundColor Yellow
Write-Host "------------------------------------------"
Write-Host "In your Automation Account, go to Modules > Browse Gallery and import:"
Write-Host "• Microsoft.Graph.Authentication"
Write-Host "• Microsoft.Graph.Applications" 
Write-Host "• Microsoft.Graph.Sites"
Write-Host ""
Write-Host "Note: Import them in this order and wait for each to complete before importing the next."
Write-Host ""

Write-Host "STEP 4: Create Automation Variables (Optional)" -ForegroundColor Yellow
Write-Host "---------------------------------------------"
Write-Host "Go to Automation Account > Variables and create:"
Write-Host "• Name: SharePointSiteUrl"
Write-Host "• Type: String"
Write-Host "• Value: https://yourtenant.sharepoint.com/sites/YourSite"
Write-Host "• Encrypted: No"
Write-Host ""

Write-Host "STEP 5: Create SharePoint List" -ForegroundColor Yellow
Write-Host "------------------------------"
Write-Host "Create a SharePoint list named 'AzureAppExpiry' with these columns:"
Write-Host "• Title (Single line of text) - for App Name"
Write-Host "• AppId (Single line of text) - for App ID"  
Write-Host "• ExpiryDate (Date and Time) - for Expiry Date"
Write-Host "• CredentialType (Choice: Secret, Certificate) - for Credential Type"
Write-Host "• KeyId (Single line of text) - for Key ID"
Write-Host ""

Write-Host "STEP 6: Import and Test the Runbook" -ForegroundColor Yellow
Write-Host "------------------------------------"
Write-Host "1. Go to Automation Account > Runbooks > Create a runbook"
Write-Host "2. Name: Get-MgAppExpiring"
Write-Host "3. Type: PowerShell"
Write-Host "4. Runtime version: 5.1"
Write-Host "5. Copy the content from 'Get-MgAppExpiring-AzureAutomation.ps1'"
Write-Host "6. Save and Publish the runbook"
Write-Host "7. Test the runbook with default parameters"
Write-Host ""

Write-Host "STEP 7: Schedule the Runbook (Optional)" -ForegroundColor Yellow
Write-Host "---------------------------------------"
Write-Host "1. Go to the runbook > Schedules > Add a schedule"
Write-Host "2. Create a new schedule (e.g., daily at 9 AM)"
Write-Host "3. Set parameters if needed:"
Write-Host "   • DaysThreshold: 30 (default)"
Write-Host "   • SharePointSiteUrl: (if not using Automation Variable)"
Write-Host "   • ListName: AzureAppExpiry (default)"
Write-Host ""

Write-Host "STEP 8: Monitor and Troubleshoot" -ForegroundColor Yellow
Write-Host "--------------------------------"
Write-Host "• Check runbook execution in Automation Account > Jobs"
Write-Host "• Review output and error logs"
Write-Host "• Verify SharePoint list is being updated"
Write-Host "• Check Managed Identity permissions if authentication fails"
Write-Host ""

Write-Host "=== Setup Complete! ===" -ForegroundColor Green
Write-Host ""
Write-Host "The runbook will:" -ForegroundColor Cyan
Write-Host "• Check all app registrations for expiring secrets and certificates"
Write-Host "• Update your SharePoint list with items expiring within the threshold"
Write-Host "• Provide detailed logging and error handling"
Write-Host "• Use Managed Identity for secure, passwordless authentication"
Write-Host ""

Write-Host "Files created:" -ForegroundColor Green
Write-Host "• Get-MgAppExpiring-AzureAutomation.ps1 (Main runbook)"
Write-Host "• Setup-AzureAutomation-AppExpiry.ps1 (This setup guide)"
Write-Host ""

Write-Host "For support or questions, refer to the Azure Automation and Microsoft Graph documentation." -ForegroundColor Gray