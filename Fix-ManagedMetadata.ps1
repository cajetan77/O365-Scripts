<#
.SYNOPSIS
Tags Managed Metadata Tags to Koru Folder
.DESCRIPTION
This script connects to  reads a CSV file containing Site Url ,Library, FOlder Path, checks whether Koru Folder content type exist in the library, if not it logs it as an error ,checks whether the folder  is tagged to Koru Folder , if not tags it with Koru Folder and associates the Managed Metadata based on the Fields
#>

param(

    [string]$CertificateThumbprint = "F4CC5E6DF6A53A34414AEE21EF66471450D4ECBE",
    [string]$ApplicationId = "054b783c-84cd-4f4c-8c92-ca35a6828679",
    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",    
    [string]$csv = ".\AllKoruSitesMetadata - Test.csv",
    [string]$TargetContentTypeName = "Koru Folder",
    [string]$siteUrl = "https://caje77sharepoint.sharepoint.com/sites/PL-InternalGovernance/"

) 


function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO",

        [string]$LogFile = "$PSScriptRoot\script_FixManagedMetadata.log"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"

    # Write to console
    Write-Host $logEntry

    # Append to log file
    Add-Content -Path $LogFile -Value $logEntry
}

Connect-PnPOnline -Url $siteUrl -ClientId $ApplicationId -Tenant $TenantId -Thumbprint $CertificateThumbprint


foreach ($site in (Import-Csv -Path $csv)) {
    #$siteUrl = $site."Site Url"
    $ListName = $site.Library
    $folderPath = $site.FolderPath
    $columnName = $site."Sharepoint Column Name"
    $Terms = $site.Terms
    $termGroup = $site."Term Group"
    $termSet = $site."Term Set"
    $managedMetadataValue = ($Terms -split ";").Trim()
    

    
        
    $list = Get-PnPList -Identity $ListName  -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Log -Message ("List '$ListName' not found in site '$siteUrl'.") -Level ERROR
        continue
    }

    $folderPath = $list.RootFolder.ServerRelativeUrl + "/" + $folderPath

    

    $cnt = Get-PnPContentType -List $list -Identity $TargetContentTypeName  -ErrorAction SilentlyContinue

    if ( -not $cnt) {
        Write-Log -Message ("Content Type '$TargetContentTypeName' not found in list '$ListName' in site '$siteUrl'.") -Level ERROR
        continue   
    }




    $termIdList = @()
    foreach ($term in $managedMetadataValue) {
        try {
            $termObject = Get-PnPTerm -Termset $termSet -TermGroup $termGroup -Identity $term
        }
        catch {
            Write-Log -Message ("Error retrieving term '$term' from TermSet '$termSet' and TermGroup '$termGroup' in site '$siteUrl' " + $_.Exception.Message) -Level ERROR
        }
        
        if (-not $termObject) {
            Write-Log -Message ("Term '$term' not found in TermSet '$termSet' and TermGroup '$termGroup' in site '$siteUrl'.") -Level ERROR
            continue
        }
        $termIdList += $termobject.Name + "|" + $termObject.Id
    }

    try {
        $folderItem = Get-PnPListItem -List $ListName -PageSize 500  | Where-Object { $_.FieldValues.FileRef -eq $folderPath } 

    }
    catch {
        Write-Log -Message ("Error getting folder item '$folderPath' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
    }
    if ( -not $folderItem) {
        Write-Log -Message ("Folder '$folderPath' not found in list '$ListName' in site '$siteUrl'.") -Level ERROR
    }
    else {
        <# Action when all if and elseif conditions are false #>
        
        foreach ($item in $folderItem) {
            try {

                $ctlistItem = Get-PnPListItem -List $ListName -Id $item.Id -IncludeContentType

                if ( $ctlistItem.ContentType.Name -ne $TargetContentTypeName ) {

                    Set-PnPListItem -List $ListName -Identity $item.Id -ContentType $TargetContentTypeName -UpdateType "SystemUpdate" -Force

                    Write-Log -Message (" Item '$($item.FieldValues.FileRef)' in list '$ListName' in '$siteUrl' has been updated successfully with content type '$TargetContentTypeName'.") -Level INFO
                }
                else {
                    Write-Log -Message (" Item '$($item.FieldValues.FileRef)' in list '$ListName' in '$siteUrl' already has the target content type '$TargetContentTypeName'. Skipping update.") -Level INFO
                }
            }
            catch {
                Write-Log -Message ("Error updating item '$($item.FieldValues.FileRef)' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
                continue
            }
    
        }
        try {

            Set-PnPListItem -List $ListName -Identity $folderItem.Id -Values @{$columnName = $termIdList } -UpdateType "SystemUpdate" -Force | Out-Null
            Write-Log -Message ("Managed metadata for folder '$folderPath' in list '$ListName' in '$siteUrl' has been set successfully.") -Level INFO
        }
        catch {
            Write-Log -Message ("Error setting managed metadata for folder '$folderPath' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
            continue 
        }

        
    }
}