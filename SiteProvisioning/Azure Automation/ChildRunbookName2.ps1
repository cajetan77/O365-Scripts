param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$ListItemId = $SiteUrl.Split('/')[-1]

$pageContentTypeName = "Doogle content category page"
$contentTypeList = @("Site Pages")
$viewsList = @("Documents", "Site Pages")
# Display name -> site column internal name (matches view field names).
$siteColumnInternalNames = @{
    "Main Category"       = "DoogleWFMainCategory"
    "Review Date"         = "ReviewDate1"
    "Notification Sent"   = "MSDNotificationSent"
    "Sub Category"        = "DoogleWFSubCategory"
    "Restricted Approval" = "DoogleWFRestrictedApproval"
}
$viewFieldInternalNames = @(
    "DocIcon", "Title", "Modified", "Editor",
    "ReviewDate1", "DoogleWFMainCategory", "MSDNotificationSent",
    "DoogleWFRestrictedApproval", "DoogleWFSubCategory"
)


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

function Publish-ContentTypesFromHub {
    param(
        [string[]]$ContentTypeNames
    )
    Write-LogMessage "Publishing content types from the Content Type Hub: $($ContentTypeNames -join ', ')" "Info"
    $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
    Write-LogMessage "Content Type Hub URL: $contentTypeHubUrl" "Info"

    try {
        Connect-PnPOnline -Url $contentTypeHubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
        $ctHub = Get-PnPContentType -ErrorAction Stop
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-LogMessage "Disconnected from content type hub" "Success"
    }
    catch {
        Write-LogMessage "Error connecting to Content Type Hub: $($_.Exception.Message)" "Error"
        throw
    }

    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop

    foreach ($cts in $ctHub) {
        if ($ContentTypeNames -contains $cts.Name) {
            Add-PnPContentTypesFromContentTypeHub -ContentTypes $cts.Id -ErrorAction Stop
            Write-LogMessage "Published content type '$($cts.Name)' to site: $SiteUrl" "Success"
        }
    }
}

function Wait-ForSiteField {
    param(
        [string]$FieldDisplayName,
        [string]$FieldInternalName,
        [int]$MaxAttempts = 12,
        [int]$DelaySeconds = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $field = $null
        if ($FieldInternalName) {
            $field = Get-PnPField -Identity $FieldInternalName -ErrorAction SilentlyContinue
        }
        if (-not $field) {
            $field = @(Get-PnPField -Identity $FieldDisplayName -ErrorAction SilentlyContinue) | Select-Object -First 1
        }
        if ($field) {
            Write-LogMessage "Site column '$FieldDisplayName' ($FieldInternalName) is available" "Success"
            return
        }

        Write-LogMessage "Waiting for site column '$FieldDisplayName' from content type hub (attempt $attempt/$MaxAttempts)..." "Info"
        Start-Sleep -Seconds $DelaySeconds
    }

    throw "Site column '$FieldDisplayName' was not provisioned after $MaxAttempts attempts."
}

function Wait-ForListField {
    param(
        [string]$ListName,
        [string]$FieldInternalName,
        [int]$MaxAttempts = 12,
        [int]$DelaySeconds = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $field = Get-PnPField -List $ListName -Identity $FieldInternalName -ErrorAction SilentlyContinue
        if ($field) {
            Write-LogMessage "Field '$FieldInternalName' is available on '$ListName'" "Success"
            return $field
        }

        Write-LogMessage "Waiting for field '$FieldInternalName' on '$ListName' (attempt $attempt/$MaxAttempts)..." "Info"
        Start-Sleep -Seconds $DelaySeconds
    }

    throw "Field '$FieldInternalName' was not provisioned on '$ListName' after $MaxAttempts attempts."
}

function Test-FieldOnList {
    param(
        $List,
        [string]$DisplayName,
        [string]$InternalName
    )

    if (Get-PnPField -List $List -Identity $DisplayName -ErrorAction SilentlyContinue) {
        return $true
    }
    if ($InternalName -and (Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue)) {
        return $true
    }
    return $false
}

function Set-ListFieldDefaultFormula {
    param(
        $List,
        [string]$FieldInternalName,
        [string]$DefaultFormula
    )

    try {
        $listField = Get-PnPField -List $List -Identity $FieldInternalName -ErrorAction Stop
        $listField.DefaultFormula = $DefaultFormula
        $listField.Update()
        Invoke-PnPQuery
        Write-LogMessage "Default formula set on '$FieldInternalName' in Documents library" "Success"
    }
    catch {
        Write-LogMessage "Could not set default formula on '$FieldInternalName': $($_.Exception.Message)" "Warning"
    }
}

function Add-ContentTypes {
    param(
        [string]$SiteUrl
    )
    try {
        $documentsLibrary = Get-PnPList -Identity "Documents" -ErrorAction Stop
        $documentsLibrary.ContentTypesEnabled = $true
        $documentsLibrary.Update()
        Invoke-PnPQuery
        Write-LogMessage "Content types enabled on $($documentsLibrary.Title)" "Success"

        # Page content type (with ReviewDate1) is applied to Site Pages only, not Documents.
        Publish-ContentTypesFromHub -ContentTypeNames @($pageContentTypeName)

        $contentType = Get-PnPContentType -Identity $pageContentTypeName -ErrorAction Stop
        foreach ($listName in $contentTypeList) {
            Add-PnPContentTypeToList -List $listName -ContentType $contentType -ErrorAction Stop
            Write-LogMessage "Added content type '$($contentType.Name)' to list '$listName'" "Success"
            Set-DefaultContentType -SiteUrl $SiteUrl -ListName $listName -ContentTypeName $pageContentTypeName
        }
    }
    catch {
        Write-LogMessage "ERROR: Failed to add content type: $($_.Exception.Message)" "Error"
        throw
    }
}

function Set-DefaultContentType {
    param(
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ContentTypeName
    )

    try {
        Set-PnPDefaultContentTypeToList -List $ListName -ContentType $ContentTypeName -ErrorAction Stop
        Write-LogMessage "Default content type set to '$ContentTypeName' on list '$ListName' ($SiteUrl)" "Success"
    }
    catch {
        Write-LogMessage "ERROR: Failed to set default content type '$ContentTypeName' on list '$ListName' ($SiteUrl): $($_.Exception.Message)" "Error"
        throw
    }
}


function Add-SiteColumns {
    param (
        [string] $siteUrl
    )
    try {
        $list = Get-PnPList -Identity "Documents"
        # These site columns originate from the page content type hub publish but must be added to Documents manually.
        $columnNames = @("Main Category", "Review Date", "Notification Sent", "Sub Category", "Restricted Approval")

        foreach ($ColumnName in $columnNames) {
            $internalName = $siteColumnInternalNames[$ColumnName]
            if ([string]::IsNullOrWhiteSpace($internalName)) {
                throw "No internal name mapped for site column '$ColumnName'."
            }

            Wait-ForSiteField -FieldDisplayName $ColumnName -FieldInternalName $internalName

            if (Test-FieldOnList -List $list -DisplayName $ColumnName -InternalName $internalName) {
                Write-LogMessage "Site column '$ColumnName' already exists in Documents library on $($siteUrl). Skipping." "Warning"
                continue
            }

            # Add existing site column to the list by internal name (SchemaXml is not returned in Azure Automation PnP).
            Add-PnPField -List $list -Field $internalName -ErrorAction Stop
            Write-LogMessage "Added site column '$ColumnName' ($internalName) to Documents library on $($siteUrl)." "Success"

            if ($ColumnName -eq "Review Date") {
                Set-ListFieldDefaultFormula -List $list -FieldInternalName $internalName -DefaultFormula "=TODAY()+365"
            }
        }

        # Views reference the internal name ReviewDate1, not the display name.
        Wait-ForListField -ListName "Documents" -FieldInternalName "ReviewDate1"
    }
    catch {
        Write-LogMessage "ERROR: Failed to add site column on $($siteUrl): $($_.Exception.Message)" "Error"
        throw
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
                $view = Get-PnPView -List $library | Where-Object { $_.DefaultView -eq $true }
                $fieldsForView = @()
                foreach ($fieldName in $viewFieldInternalNames) {
                    $field = Get-PnPField -List $library -Identity $fieldName -ErrorAction SilentlyContinue
                    if ($field) {
                        $fieldsForView += $fieldName
                    }
                    else {
                        Write-LogMessage "Field '$fieldName' not on $($library.Title); omitting from view" "Warning"
                    }
                }

                if ($fieldsForView.Count -eq 0) {
                    throw "No view fields are available on list '$list'."
                }

                Set-PnPView -List $library -Identity $view.Id -Fields $fieldsForView -ErrorAction Stop
                Write-LogMessage "Custom view updated on $($library.Title) with fields: $($fieldsForView -join ', ')" "Success"
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
