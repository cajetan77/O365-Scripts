<#
.SYNOPSIS
Tags Managed Metadata Tags to Koru Folder using simple PnP batching
.DESCRIPTION
This script uses PnP batch operations to improve performance while keeping the original logic simple
#>

param(
    [string]$CertificateThumbprint = "F4CC5E6DF6A53A34414AEE21EF66471450D4ECBE",
    [string]$ApplicationId = "054b783c-84cd-4f4c-8c92-ca35a6828679",
    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",    
    [string]$csv = ".\AllKoruSitesMetadata - Test.csv",
    [string]$TargetContentTypeName = "Koru Folder",
    [string]$siteUrl = "https://caje77sharepoint.sharepoint.com/sites/PL-InternalGovernance/",
    [int]$BatchSize = 100
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

# Load CSV data
$csvData = Import-Csv -Path $csv
Write-Log "Processing $($csvData.Count) items with batch size of $BatchSize"

# Create batch and counter
$batch = New-PnPBatch
$batchCount = 0

foreach ($site in $csvData) {
    $ListName = $site.Library
    $folderPath = $site.FolderPath
    $columnName = $site."Sharepoint Column Name"
    $Terms = $site.Terms
    $termGroup = $site."Term Group"
    $termSet = $site."Term Set"
    $managedMetadataValue = ($Terms -split ";").Trim()
    
    # Get list
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Log -Message ("List '$ListName' not found in site '$siteUrl'.") -Level ERROR
        continue
    }

    $folderPath = $list.RootFolder.ServerRelativeUrl + "/" + $folderPath

    # Check content type exists
    $cnt = Get-PnPContentType -List $list -Identity $TargetContentTypeName -ErrorAction SilentlyContinue
    if (-not $cnt) {
        Write-Log -Message ("Content Type '$TargetContentTypeName' not found in list '$ListName' in site '$siteUrl'.") -Level ERROR
        continue   
    }

    # Get terms
    $termIdList = @()
    foreach ($term in $managedMetadataValue) {
        try {
            $termObject = Get-PnPTerm -Termset $termSet -TermGroup $termGroup -Identity $term
            if ($termObject) {
                $termIdList += $termobject.Name + "|" + $termObject.Id
            }
        }
        catch {
            Write-Log -Message ("Error retrieving term '$term' from TermSet '$termSet' and TermGroup '$termGroup' in site '$siteUrl' " + $_.Exception.Message) -Level ERROR
        }
    }

    # Get folder item
    try {
        $folderItem = Get-PnPListItem -List $ListName -PageSize 500 | Where-Object { $_.FieldValues.FileRef -eq $folderPath } 
    }
    catch {
        Write-Log -Message ("Error getting folder item '$folderPath' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
        continue
    }

    if (-not $folderItem) {
        Write-Log -Message ("Folder '$folderPath' not found in list '$ListName' in site '$siteUrl'.") -Level ERROR
        continue
    }
    
    # Process folder items
    foreach ($item in $folderItem) {
        try {
            $ctlistItem = Get-PnPListItem -List $ListName -Id $item.Id -IncludeContentType

            # Add content type update to batch if needed
            if ($ctlistItem.ContentType.Name -ne $TargetContentTypeName) {
                Set-PnPListItem -List $ListName -Identity $item.Id -ContentType $TargetContentTypeName -UpdateType "SystemUpdate" -Force -Batch $batch
                $batchCount++
                Write-Log -Message ("Added content type update to batch for item '$($item.FieldValues.FileRef)'") -Level DEBUG
            }
            else {
                Write-Log -Message ("Item '$($item.FieldValues.FileRef)' already has the target content type '$TargetContentTypeName'. Skipping update.") -Level INFO
            }
        }
        catch {
            Write-Log -Message ("Error preparing content type update for item '$($item.FieldValues.FileRef)' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
        }
    }

    # Add metadata update to batch
    try {
        Set-PnPListItem -List $ListName -Identity $folderItem.Id -Values @{$columnName = $termIdList } -UpdateType "SystemUpdate" -Force -Batch $batch
        $batchCount++
        Write-Log -Message ("Added metadata update to batch for folder '$folderPath'") -Level DEBUG
    }
    catch {
        Write-Log -Message ("Error preparing metadata update for folder '$folderPath' in list '$ListName' in '$siteUrl' " + $_.Exception.Message) -Level ERROR
    }

    # Execute batch when it reaches batch size
    if ($batchCount -ge $BatchSize) {
        try {
            Write-Log -Message ("Executing batch with $batchCount operations") -Level INFO
            Invoke-PnPBatch -Batch $batch
            Write-Log -Message ("Batch executed successfully") -Level INFO
            
            # Create new batch
            $batch = New-PnPBatch
            $batchCount = 0
        }
        catch {
            Write-Log -Message ("Batch execution failed: " + $_.Exception.Message) -Level ERROR
            
            # Create new batch to continue
            $batch = New-PnPBatch
            $batchCount = 0
        }
    }
}

# Execute remaining batch operations
if ($batchCount -gt 0) {
    try {
        Write-Log -Message ("Executing final batch with $batchCount operations") -Level INFO
        Invoke-PnPBatch -Batch $batch
        Write-Log -Message ("Final batch executed successfully") -Level INFO
    }
    catch {
        Write-Log -Message ("Final batch execution failed: " + $_.Exception.Message) -Level ERROR
    }
}

Write-Log -Message ("Script completed") -Level INFO