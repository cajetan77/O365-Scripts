param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$ListItemId = $SiteUrl.Split('/')[-1]


function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $idPrefix = if ($script:ListItemId) { "[ItemId:$($script:ListItemId)] " } else { "" }
    $logMessage = "[$timestamp] [$Level] $idPrefix$Message"
    
    # Azure Automation compatible logging
    switch ($Level) {
        "Error" { 
            Write-Error $logMessage
        }
        "Warning" { 
            Write-Warning $logMessage
        }
        default { 
            Write-Output $logMessage
        }
    }
}

Write-LogMessage "ListItemId: $ListItemId" "Info"
try {
    $tenantId = Get-AutomationVariable -Name "TenantId" -ErrorAction Stop
    $ClientId = Get-AutomationVariable -Name "AppId" -ErrorAction Stop
    $TenantName = Get-AutomationVariable -Name "TenantName" -ErrorAction Stop
    $KeyVaultName = Get-AutomationVariable -Name "KeyVaultName" -ErrorAction Stop
    #  $CertificateName = Get-AutomationVariable -Name "CertificateName" -ErrorAction Stop
    $Thumbprint = Get-AutomationVariable -Name "Thumbprint" -ErrorAction Stop
    $ResourceGroupName = Get-AutomationVariable -Name "ResourceGroupName" -ErrorAction Stop
    $AutomationAccountName = Get-AutomationVariable -Name "AutomationAccountName" -ErrorAction Stop
}
catch {
    Write-LogMessage "[ItemId:$ListItemId] Failed to retrieve required Automation Variables: $($_.Exception.Message)" "Error"
    throw
}


try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
}
catch {
    Write-LogMessage "Failed to connect to SharePoint: $($_.Exception.Message)" "Error"
    throw
}


function Set-SiteRegionalSettings {
    param(
        [string]$SiteUrl,
        [string]$siteName 
    )
 
    try {
        $web = Get-PnpWeb -Includes RegionalSettings.LocaleId, RegionalSettings.TimeZones 
        $localeId = 5129 # New Zealand English
        $web.RegionalSettings.LocaleId = $localeId
        $web.Update()
        Invoke-PnPQuery
        Write-LogMessage "Updated Site Regional Settings to have NZ Time Zone and NZ Locale  $($web.Url)" "Info"
    }
    catch {
        Write-LogMessage "Error connecting to site $SiteUrl :$($_.Exception.Message)" "Error" 
    
    }
    

}  

function Add-GroupstoSharePointGroups {
    param (
        [string]$SiteUrl,
        [string]$siteName
    )

    try {
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $EntraGroupObjectId = "b4fa1e98-2893-4f26-bc64-f7a3e93b3753" # Replace with actual Entra Group Object ID
        $groupLoginName = "c:0t.c|tenant|$EntraGroupObjectId"
        Add-PnPGroupMember `
            -Group $ownersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-LogMessage "Added Entra group with Object ID $EntraGroupObjectId to Owners group on $($SiteUrl)" "Success"
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop
        Add-PnPGroupMember `
            -Group $membersGroup.Title `
            -LoginName $groupLoginName `
            -ErrorAction Stop
        Write-LogMessage "Added Entra group with Object ID $EntraGroupObjectId to Members group on $($SiteUrl)" "Success"
    
    }
    catch {
        Write-LogMessage "ERROR: Failed to add group to SharePoint group on $($SiteUrl): $($_.Exception.Message)" "Error"
    }
    
}


function Set-SearchSettings {
    param (
        [string]$SiteUrl
    )
    try {

        $list = Get-PnPList -Identity "Site Assets" -ErrorAction SilentlyContinue
        if ($list) {
            $list.NoCrawl = $true
            $list.Update()
            Invoke-PnPQuery
            Write-LogMessage "Site Assets list no crawled on $($SiteUrl)" "Success"
        }
        else {
            $web = Get-PnPWeb 
            $web.Lists.EnsureSiteAssetsLibrary()
            Invoke-PnPQuery
            $list = Get-PnPList -Identity "Site Assets" 
            Set-PnPList -Identity $list -NoCrawl:$true
            Invoke-PnPQuery
            Write-LogMessage "Site Assets list no crawled on $($SiteUrl)" "Success"

        }

    }
    catch {
        Write-LogMessage "ERROR: Failed to set search settings on $($SiteUrl): $($_.Exception.Message)" "Error"
       
    }
}


function Install-App {
    param(
        [string]$SiteUrl
    )
    try {
        $appIds = @("59903278-DD5D-4E9E-BEF6-562AAE716B8B", "00406271-0276-406F-9666-512623EB6709")
      
        foreach ($appId in $appIds) {
            $app = Get-PnPApp -Identity $appId  -ErrorAction Stop
            if ($null -eq $app.InstalledVersion) {
                Install-PnPApp -Identity $app  -ErrorAction Stop
                Write-LogMessage "App installed: $($app.Title)" "Success"    
            }
            else {
                Write-LogMessage "App already installed: $($app.Title)" "Warning"
            }
           
        }
    }
    catch {
        Write-LogMessage "ERROR: Failed to install app on $($SiteUrl): $($_.Exception.Message)" "Error"
    }
}


function SavePageTemplate {
    param (
        [string] $siteUrl
    )
    try {
        Connect-AzAccount -Identity -Force
        $StorageAccountName = "spostoragecaj134"
        $ContainerName = "images"
        #$BlobName = "pexels-padrinan-255379.jpg"
        $BlobTemplateXML = "SiteProvisioningtemplate.xml"

        $DestinationPath = "$env:TEMP\pexels-padrinan-255379.jpg"
        $LocalTemplateXML = "$env:TEMP\SiteProvisioningtemplate.xml"

        $Context = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
        
        
        Write-LogMessage "Downloading site template configuration from Azure Storage..." "Info"
        Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobTemplateXML -Destination $LocalTemplateXML -Context $Context -Force

        # 9. Execute the Template (Fixed Target Variable Name Mismatch)
        Write-LogMessage "Executing Invoke-PnPSiteTemplate against target web..." "Info"
        Invoke-PnPSiteTemplate -Path $LocalTemplateXML

        Write-LogMessage "Page template added: $($LocalTemplateXML)" "Success"
        
        #Save-PnPSiteTemplate -Template $pagesTemplate -Out ("{0}{1}.xml" -f $saveTemplateLocation, $templateName)
    }
    catch {
        Write-LogMessage "ERROR: Failed to save page template on $($siteUrl): $($_.Exception.Message)" "Error"
    }

   
}

Function Set-DocLibraryPermissions {
    param(
        [string]$SiteUrl
    )
    try {
        Write-LogMessage "Setting DocLibraryPermissions on $SiteUrl" "Info"
        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-LogMessage "ERROR: $($library.Title) not found on $SiteUrl" "Error"
           
        }

        Set-PnPList -Identity $library -BreakRoleInheritance -CopyRoleAssignments -ErrorAction Stop

        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop

        if (-not $ownersGroup -or -not $membersGroup) {
            Write-LogMessage -Message "ERROR: Could not resolve Owners or Members group for $SiteUrl." -Level Error -siteName $siteName
        }

        # Ensure both site groups have Contribute on Documents.
        Set-PnPListPermission `
            -Identity $library `
            -Group $ownersGroup.Title `
            -RemoveRole "Full Control" `
            -AddRole "Contribute" `
            
        Set-PnPListPermission `
            -Identity $library `
            -Group $membersGroup.Title `
            -RemoveRole "Edit" `
            -AddRole "Contribute" `
            
        Write-LogMessage "Permissions updated: Owners and Members have Contribute on Documents." "Success"
    }
    catch {
        Write-LogMessage "ERROR: Failed to set DocLibraryPermissions: $($_.Exception.Message)" "Error"
       
    }

}


function Set-Branding {
    param (
        [string]$SiteUrl`
        
    )
    try {

        Connect-AzAccount -Identity -Force
        $StorageAccountName = "spostoragecaj134"
        $ContainerName = "images"
        $BlobName = "pexels-padrinan-255379.jpg"
        $DestinationPath = "$env:TEMP\pexels-padrinan-255379.jpg"
        try {
            $Context = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount
            Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -Destination $DestinationPath -Context $Context -Force   
        }
        catch {
            Write-LogMessage "ERROR: Failed to get blob from Azure Storage: $($_.Exception.Message)" "Error"
            throw
        }
        # 1. Make sure layout is Extended
        Set-PnPWebHeader -HeaderLayout Extended

        $SPFile = Add-PnPFile -Path $DestinationPath -Folder "SiteAssets"

        # 4. Apply in one call
        Set-PnPWebHeader `
            -HeaderLayout Extended `
            -HeaderBackgroundImageUrl $url `
            -HeaderBackgroundImageFocalX 0.5 `
            -HeaderBackgroundImageFocalY 0.5
        #$bgUrl = "/sites/WRK-Test4/SiteAssets/pexels-padrinan-255379.jpg"

        Set-PnPWebHeader -HeaderBackgroundImageUrl $bgUrl -HeaderLayout Extended  -ErrorAction Stop
        Set-PnPWebHeader -HeaderLayout "Extended"
        Set-PnPFooter -Layout "Extended"
        Set-PnPWeb -MegaMenuEnabled:$false
        Set-PnPWeb -SiteLogoUrl $SPFile.ServerRelativeUrl
        Set-PnPWebHeader -SiteThumbnailUrl $SPFile.ServerRelativeUrl
       
      
        Write-LogMessage "Header and Footer Extended  on $($SiteUrl)" "Success"
    }
    catch {
        Write-LogMessage "ERROR: Failed to set branding on $($SiteUrl): $($_.Exception.Message)" "Error"
        
    }
}



Set-SiteRegionalSettings -SiteUrl $SiteUrl 
Add-GroupstoSharePointGroups -SiteUrl $SiteUrl 
Set-SearchSettings -SiteUrl $SiteUrl 
Install-App -SiteUrl $SiteUrl 
Set-DocLibraryPermissions -SiteUrl $SiteUrl 
Set-Branding -SiteUrl $SiteUrl 

$TargetParameters = @{
    SiteUrl = $SiteUrl
}
try {
    Start-AzAutomationRunbook -AutomationAccountName $AutomationAccountName -Name "ChildRunbookName2" -ResourceGroupName $ResourceGroupName -Parameters $TargetParameters
    Write-LogMessage "Runbook started" "Success"
}
catch {
    Write-LogMessage "ERROR: Failed to start runbook: $($_.Exception.Message)" "Error"
}

if(Get-PnPContext) {
    Disconnect-PnPOnline
    Write-LogMessage "Disconnected from SharePoint" "Info"
}
if(Get-AzContext) {
    Disconnect-AzAccount
    Write-LogMessage "Disconnected from Azure" "Info"
}


