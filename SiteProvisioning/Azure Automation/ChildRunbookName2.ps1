param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$ListItemId = $SiteUrl.Split('/')[-1]

$contentTypeName = "content category page"
$contentTypeList = @("Site Pages")
$viewsList = @("Documents", "Site Pages")


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

function Get-ContentTypeHub {
    param(
        [string]$ct,
        [string]$siteName
    )
    Write-LogMessage "Adding Content Type from the Content Type Hub" "Info"
    $contentTypesArray = $ct.Split(", ") | ForEach-Object { $_.Trim() }  
    $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
    Write-LogMessage "Content Type Hub URL: $contentTypeHubUrl" "Info"
    try {
        $ctconnection = Connect-PnPOnline -Url $contentTypeHubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
     
        $ctHub = Get-PnPContentType -Connection $ctconnection
        Disconnect-PnPOnline
        Write-LogMessage "Disconnected from content type hub" "Success"
    }
    
    catch {
        Write-LogMessage "Error connecting to Content Type Hub: $($_.Exception.Message)" "Error"
    }
 
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
       
        
        foreach ($cts in $ctHub) {
            if ($contentTypesArray -contains $cts.Name) {
                Add-PnPContentTypesFromContentTypeHub -ContentTypes $cts.Id -Site $SiteUrl -Connection $siteconnection
                Write-LogMessage -Message "Added content type '$($cts.Name)' to site: $SiteUrl" -Level Success -siteName $siteName
            }
        }
     
    }
    catch {
        Write-LogMessage -Message "Error adding Content Types from Hub: $($_.Exception.Message)" -Level Error -siteName $siteName
    }
    
}


function Add-ContentTypes {
    param(
        [string]$SiteUrl,
        [string]$siteName
    )
    try {

        $library = Get-PnPList -Identity "Documents" -ErrorAction Stop
        if (-not $library) {
            Write-LogMessage "ERROR: DocLibrary not found on $SiteUrl" "Error"
            
            
        }
        else {
            $library.ContentTypesEnabled = $true
            $library.Update()
            Invoke-PnPQuery
            Write-LogMessage "Content types enabled on $($library.Title)" "Success"
            Get-ContentTypeHub -ct $contentTypeName  -siteName $siteName
            #Get the content type
            $ContentType = Get-PnPContentType -Identity $contentTypeName
            If ($ContentType) {
                #Add Content Type to Library
                foreach ($listName in $contentTypeList) {  
                    Add-PnPContentTypeToList -List $listName -ContentType $ContentType
                   
                    Write-LogMessage "Added content type '$($ContentType.Name)' to list '$($listName)'" "Success"
                    Set-DefaultContentType -SiteUrl $SiteUrl -ListName $listName -ContentTypeName $contentTypeName -siteName $siteName
                }
            }
        }
    }
    catch {
        Write-LogMessage "ERROR: Failed to add content type: $($_.Exception.Message)" "Error"
        
    }
}

function Set-DefaultContentType {
    param(
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ContentTypeName,
        [string]$siteName
    )

    try {
        Set-PnPDefaultContentTypeToList -List $ListName -ContentType $ContentTypeName -ErrorAction Stop
        Write-LogMessage "Default content type set to '$ContentTypeName' on list '$ListName' ($SiteUrl)" "Success"
    }
    catch {
        Write-LogMessage "ERROR: Failed to set default content type '$ContentTypeName' on list '$ListName' ($SiteUrl): $($_.Exception.Message)" "Error"
        
    }
}


function Add-SiteColumns {
    param (
        [string] $siteUrl
    )
    try {
        $list = Get-PnPList -Identity "Documents"
        $columnNames = @("Main Category", "Review Date", "Notification Sent", "Sub Category", "Restricted Approval")
        foreach ($ColumnName in $columnNames) {
            $existingColumn = Get-PnPField -Identity $ColumnName -ErrorAction SilentlyContinue
            if ($existingColumn) {

                Write-LogMessage "Site column '$ColumnName'  exists on $($siteUrl). Skipping creation." "Warning"
                switch ($columnName) {
                    "Main Category" { 
                        #Write-Host "Site column '$ColumnName' already exists on $($siteUrl). Skipping adding to list." -ForegroundColor Yellow
                        #Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                        #Write-Host "Added site column '$columnName' to '$listTitle'"
                    }
                    "Review Date" {
                        $existingColumn.DefaultFormula = "=TODAY()+365"
                        $existingColumn.UpdateAndPushChanges($true)
                        Invoke-PnPQuery
                    }
                }
                $fieldInList = Get-PnPField -List $list -Identity $ColumnName -ErrorAction SilentlyContinue
                if ($fieldInList) {
                    Write-LogMessage -Message "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." -Level Warning -siteName $siteName
                }
                else {
                    switch ($columnName) {
                        "Main Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-LogMessage "Added site column '$columnName' to '$list'" "Success"
                        }
                        "Sub Category" { 
                            Add-PnPFieldFromXml -List $list -FieldXml $existingColumn.SchemaXml -ErrorAction Stop
                            Write-LogMessage "Added site column '$columnName' to '$list'" "Success"
                        }
                        Default {
                            Add-PnPField -List $list -Field $existingColumn
                            Write-LogMessage "Added site column '$columnName' to '$list'" "Success"
                        }
                    }
                    
                    Write-LogMessage "Added existing site column '$ColumnName' to Documents library on $($siteUrl)." "Success"
                }
            } 
            else {
                Write-LogMessage "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." "Warning"
            }

        }

    }
    catch {
        Write-LogMessage "ERROR: Failed to add site column on $($siteUrl): $($_.Exception.Message)" "Error"
    }
    
}

function Set-Views {
    param(
        [string]$SiteUrl
    )
    try {
        
        foreach ($list in $viewsList) {
        
            $library = Get-PnPList -Identity $list -ErrorAction Stop
            if (-not $library) {
                Write-LogMessage "ERROR: DocLibrary not found on $SiteUrl" "Error"           
            }
            else {
            
                $view = Get-PnPView -List $library  | Where-Object { $_.DefaultView -eq $true }
                Set-PnPView -List $library -Identity $view.Id -Fields "DocIcon", "Title", "Modified", "Editor", "ReviewDate1", "DoogleWFMainCategory", "MSDNotificationSent", "DoogleWFRestrictedApproval", "DoogleWFSubCategory" -ErrorAction Stop
                Write-LogMessage "Custom view created and set as default on $($library.Title)" "Success"
            }
        }
    }
    catch {
        Write-LogMessage "ERROR: Failed to set views on $($SiteUrl): $($_.Exception.Message)" "Error"
        
    }
    
}

function Connect-Site {
    param(
        [string]$SiteUrl
    )
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
    }
    catch {
        Write-LogMessage "ERROR: Failed to connect to site $($SiteUrl): $($_.Exception.Message)" "Error"
    }
}

Connect-Site -SiteUrl $SiteUrl
Add-ContentTypes -SiteUrl $SiteUrl
Add-SiteColumns -siteUrl $SiteUrl
Set-Views -SiteUrl $SiteUrl
