param(
    [Parameter(Mandatory = $true)]
    [string]$RequestBody
)

# Ka huri i te kupu raw JSON hei puka ahanoa PowerShell (Object)
$Data = ConvertFrom-Json $RequestBody

# Inaianei, ka taea te rāwekeweke i ngā mara katoa i tukuna mai:
Write-Output "Project Name: $($Data.ProjectName)"
Write-Output "Custom Field 1: $($Data.ObjectUrl)"

# ===========================
# HELPER FUNCTIONS
# ===========================

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


Write-LogMessage " Starting site availability validation script" "Info"


try {
    $tenantId = Get-AutomationVariable -Name "TenantId" -ErrorAction Stop
    $ClientId = Get-AutomationVariable -Name "AppId" -ErrorAction Stop
    $TenantName = Get-AutomationVariable -Name "TenantName" -ErrorAction Stop
    $KeyVaultName = Get-AutomationVariable -Name "KeyVaultName" -ErrorAction Stop
    $Thumbprint = Get-AutomationVariable -Name "Thumbprint" -ErrorAction Stop
    #  $CertificateName = Get-AutomationVariable -Name "CertificateName" -ErrorAction Stop
    $ResourceGroupName = Get-AutomationVariable -Name "ResourceGroupName" -ErrorAction Stop
    $AutomationAccountName = Get-AutomationVariable -Name "AutomationAccountName" -ErrorAction Stop
}
catch {
    Write-LogMessage "[ItemId:$ListItemId] Failed to retrieve required Automation Variables: $($_.Exception.Message)" "Error"
    throw
}

$SiteUrl = $Data.ObjectUrl
$AdminUrl = "https://$TenantName-admin.sharepoint.com"
Write-LogMessage "Target site URL: $SiteUrl" "Info" 



Write-LogMessage "Connecting to SharePoint Admin via PnP (app-only)..." "Info"
try {
    
    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $tenantId -Thumbprint $Thumbprint -ErrorAction Stop
    # Connect-PnPOnline -Url $AdminUrl -ManagedIdentity
    $context = Get-PnPContext
    if (-not $context) {
        throw "Failed to establish PnP context"
    }
    
    Write-LogMessage "Successfully connected to SharePoint Admin." "Success"
}
catch {
    Write-LogMessage "Failed to connect to SharePoint Admin using PnP: $($_.Exception.Message)" "Error"
    throw
}

$AdminUrl = "https://$TenantName-admin.sharepoint.com"
Write-LogMessage "Checking if site already exists..." "Info"
try {
    $SiteExists = Get-PnPTenantSite | Where-Object { $_.Url -eq $SiteUrl } -ErrorAction Stop
    
    #Write-LogMessage "Checking recycle bin for deleted site..." "Info"
    #$SiteExistsInRecycleBin = Get-PnPTenantRecycleBinItem | Where-Object { $_.Url -eq $SiteUrl } -ErrorAction Stop
    
    if ($SiteExists -or $SiteExistsInRecycleBin) {
        $siteId = if ($SiteExists) { $SiteExists.Id } else { $SiteExistsInRecycleBin.Id }
        $location = if ($SiteExists) { "active sites" } else { "recycle bin" }
        
        Write-LogMessage "Site URL $SiteUrl already exists in tenant ($location)" "Warning"
        
        $out = @{
            Status   = "Exists"
            Message  = "Site $Title already exists with Id: $siteId"
            SiteUrl  = $SiteUrl
            Location = $location
        }
        Write-Output ($out | ConvertTo-Json -Compress)
        
        try {

            $TargetParameters = @{
                SiteUrl = $SiteUrl
            }
            Connect-AzAccount -Identity
            Start-AzAutomationRunbook -AutomationAccountName $AutomationAccountName -Name "ChildRunbookName" -ResourceGroupName $ResourceGroupName -Parameters $TargetParameters
        }
        catch {
            Write-LogMessage "Error connecting to Azure: $($_.Exception.Message)" "Error"
            throw
        }
        
        Write-LogMessage "Site validation completed - site already exists" "Info"
        
    }
    else {
        Write-LogMessage "Site URL $SiteUrl is not yet created  " "Warning"
        $out = @{
            Status  = "NotExists"
            Message = "Site $Title is not yet created"
            SiteUrl = $SiteUrl
        }
        Write-Output ($out | ConvertTo-Json -Compress)
        exit 0
    }
}
catch {
    Write-LogMessage "Error checking site availability: $($_.Exception.Message)" "Error"
    throw
}
  
# ===========================
# PREPARE RUNBOOK PARAMETERS
# ===========================



# Cleanup: Disconnect from SharePoint and Azure (best practice for Azure Automation)
try {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-LogMessage "Disconnected from SharePoint Admin" "Info"
}
catch {
    # Ignore disconnect errors
}

try {
    Disconnect-AzAccount -ErrorAction SilentlyContinue
    Write-LogMessage "Disconnected from Azure" "Info"
}
catch {
    # Ignore disconnect errors
}



