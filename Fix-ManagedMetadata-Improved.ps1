<#
.SYNOPSIS
    Tags Managed Metadata Tags to Koru Folder using PnP batch operations for improved performance.

.DESCRIPTION
    This script connects to SharePoint, reads a CSV file containing Site Url, Library, Folder Path,
    checks whether Koru Folder content type exists in the library, checks whether the folder is tagged
    to Koru Folder, and associates the Managed Metadata based on the Fields using batch operations
    for better performance.

.PARAMETER CertificateThumbprint
    The certificate thumbprint for authentication

.PARAMETER ApplicationId
    The application ID for authentication

.PARAMETER TenantId
    The tenant ID for authentication

.PARAMETER CsvPath
    Path to the CSV file containing metadata to process

.PARAMETER TargetContentTypeName
    The target content type name (default: "Koru Folder")

.PARAMETER SiteUrl
    The SharePoint site URL to connect to

.PARAMETER BatchSize
    Number of operations to batch together (default: 100)

.EXAMPLE
    .\Fix-ManagedMetadata-Improved.ps1 -CsvPath ".\AllKoruSitesMetadata.csv"
#>

param(
    [string]$CertificateThumbprint = "F4CC5E6DF6A53A34414AEE21EF66471450D4ECBE",
    [string]$ApplicationId = "054b783c-84cd-4f4c-8c92-ca35a6828679",
    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",    
    [string]$CsvPath = ".\AllKoruSitesMetadata - Test.csv",
    [string]$TargetContentTypeName = "Koru Folder",
    [string]$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/PL-InternalGovernance/",
    [int]$BatchSize = 100
) 

# Global variables for tracking
$script:LogFile = "$PSScriptRoot\script_FixManagedMetadata_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:ProcessedCount = 0
$script:SuccessCount = 0
$script:ErrorCount = 0
$script:SkippedCount = 0

# Cache for performance optimization
$script:ListCache = @{}
$script:ContentTypeCache = @{}
$script:TermCache = @{}

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG", "SUCCESS")]
        [string]$Level = "INFO",
        [string]$LogFile = $script:LogFile
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"

    # Color coding for console output
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN" { "Yellow" }
        "SUCCESS" { "Green" }
        "DEBUG" { "Gray" }
        default { "White" }
    }

    Write-Host $logEntry -ForegroundColor $color
    Add-Content -Path $LogFile -Value $logEntry
}

function Initialize-Script {
    Write-Log "Starting Fix-ManagedMetadata script with batch operations" -Level INFO
    Write-Log "CSV Path: $CsvPath" -Level INFO
    Write-Log "Target Content Type: $TargetContentTypeName" -Level INFO
    Write-Log "Site URL: $SiteUrl" -Level INFO
    Write-Log "Batch Size: $BatchSize" -Level INFO
    Write-Log "Log File: $script:LogFile" -Level INFO
    
    # Validate CSV file exists
    if (-not (Test-Path $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level ERROR
        throw "CSV file not found"
    }
}

function Connect-ToSite {
    try {
        Write-Log "Connecting to SharePoint site: $SiteUrl" -Level INFO
        Connect-PnPOnline -Url $SiteUrl -ClientId $ApplicationId -Tenant $TenantId -Thumbprint $CertificateThumbprint
        Write-Log "Successfully connected to SharePoint site" -Level SUCCESS
    }
    catch {
        Write-Log "Error connecting to SharePoint site: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Get-CachedList {
    param([string]$ListName)
    
    if (-not $script:ListCache.ContainsKey($ListName)) {
        try {
            $list = Get-PnPList -Identity $ListName -ErrorAction Stop
            $script:ListCache[$ListName] = $list
            Write-Log "Cached list: $ListName" -Level DEBUG
        }
        catch {
            Write-Log "List '$ListName' not found in site '$SiteUrl'" -Level ERROR
            $script:ListCache[$ListName] = $null
        }
    }
    
    return $script:ListCache[$ListName]
}

function Get-CachedContentType {
    param([string]$ListName, [string]$ContentTypeName)
    
    $cacheKey = "$ListName|$ContentTypeName"
    
    if (-not $script:ContentTypeCache.ContainsKey($cacheKey)) {
        try {
            $list = Get-CachedList -ListName $ListName
            if ($list) {
                $contentType = Get-PnPContentType -List $list -Identity $ContentTypeName -ErrorAction Stop
                $script:ContentTypeCache[$cacheKey] = $contentType
                Write-Log "Cached content type: $ContentTypeName for list: $ListName" -Level DEBUG
            }
            else {
                $script:ContentTypeCache[$cacheKey] = $null
            }
        }
        catch {
            Write-Log "Content Type '$ContentTypeName' not found in list '$ListName'" -Level ERROR
            $script:ContentTypeCache[$cacheKey] = $null
        }
    }
    
    return $script:ContentTypeCache[$cacheKey]
}

function Get-CachedTerms {
    param([array]$ManagedMetadataValues, [string]$TermSet, [string]$TermGroup)
    
    $termIdList = @()
    
    foreach ($term in $ManagedMetadataValues) {
        $cacheKey = "$TermGroup|$TermSet|$term"
        
        if (-not $script:TermCache.ContainsKey($cacheKey)) {
            try {
                $termObject = Get-PnPTerm -Termset $TermSet -TermGroup $TermGroup -Identity $term -ErrorAction Stop
                $script:TermCache[$cacheKey] = $termObject
                Write-Log "Cached term: $term from TermSet: $TermSet" -Level DEBUG
            }
            catch {
                Write-Log "Error retrieving term '$term' from TermSet '$TermSet' and TermGroup '$TermGroup': $($_.Exception.Message)" -Level ERROR
                $script:TermCache[$cacheKey] = $null
            }
        }
        
        $termObject = $script:TermCache[$cacheKey]
        if ($termObject) {
            $termIdList += $termObject.Name + "|" + $termObject.Id
        }
        else {
            Write-Log "Term '$term' not found in TermSet '$TermSet' and TermGroup '$TermGroup'" -Level ERROR
        }
    }
    
    return $termIdList
}

function Get-FolderItems {
    param([string]$ListName, [string]$FolderPath)
    
    try {
        $list = Get-CachedList -ListName $ListName
        if (-not $list) {
            return $null
        }
        
        $fullFolderPath = $list.RootFolder.ServerRelativeUrl + "/" + $FolderPath
        
        # Get folder items efficiently
        $folderItems = Get-PnPListItem -List $ListName -PageSize 500 | 
        Where-Object { $_.FieldValues.FileRef -eq $fullFolderPath }
        
        if (-not $folderItems) {
            Write-Log "Folder '$fullFolderPath' not found in list '$ListName'" -Level ERROR
            return $null
        }
        
        return @{
            Items    = $folderItems
            FullPath = $fullFolderPath
        }
    }
    catch {
        Write-Log "Error getting folder items for '$FolderPath' in list '$ListName': $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Process-ItemsBatch {
    param(
        [array]$BatchItems,
        [int]$BatchNumber
    )
    
    Write-Log "Processing batch $BatchNumber with $($BatchItems.Count) operations" -Level INFO
    
    # Create new batch
    $batch = New-PnPBatch
    $batchOperations = @()
    
    # Group operations by type for better organization
    $contentTypeUpdates = @()
    $metadataUpdates = @()
    
    foreach ($batchItem in $BatchItems) {
        $item = $batchItem.Item
        $operation = $batchItem.Operation
        $data = $batchItem.Data
        
        try {
            switch ($operation) {
                "UpdateContentType" {
                    # Add content type update to batch
                    Set-PnPListItem -List $data.ListName -Identity $item.Id -ContentType $data.ContentTypeName -UpdateType "SystemUpdate" -Force -Batch $batch
                    $contentTypeUpdates += @{
                        Item            = $item
                        ListName        = $data.ListName
                        ContentTypeName = $data.ContentTypeName
                    }
                }
                "UpdateMetadata" {
                    # Add metadata update to batch
                    Set-PnPListItem -List $data.ListName -Identity $item.Id -Values @{$data.ColumnName = $data.TermIdList } -UpdateType "SystemUpdate" -Force -Batch $batch
                    $metadataUpdates += @{
                        Item       = $item
                        ListName   = $data.ListName
                        ColumnName = $data.ColumnName
                        FolderPath = $data.FolderPath
                    }
                }
            }
            
            $batchOperations += @{
                Operation = $operation
                Item      = $item
                Data      = $data
            }
        }
        catch {
            Write-Log "Error preparing batch operation for item $($item.FieldValues.FileRef): $($_.Exception.Message)" -Level ERROR
            $script:ErrorCount++
        }
    }
    
    # Execute batch if there are operations
    if ($batchOperations.Count -gt 0) {
        try {
            Write-Log "Executing batch with $($batchOperations.Count) operations" -Level INFO
            Invoke-PnPBatch -Batch $batch
            
            # Log success for each operation in the batch
            foreach ($operation in $batchOperations) {
                switch ($operation.Operation) {
                    "UpdateContentType" {
                        Write-Log "Content type updated for item: $($operation.Item.FieldValues.FileRef)" -Level SUCCESS
                    }
                    "UpdateMetadata" {
                        Write-Log "Managed metadata updated for folder: $($operation.Data.FolderPath)" -Level SUCCESS
                    }
                }
                $script:SuccessCount++
                $script:ProcessedCount++
            }
        }
        catch {
            Write-Log "Batch execution failed: $($_.Exception.Message)" -Level ERROR
            
            # Process items individually as fallback
            Write-Log "Falling back to individual processing for this batch" -Level WARN
            Process-ItemsIndividually -Operations $batchOperations
        }
    }
}

function Process-ItemsIndividually {
    param([array]$Operations)
    
    Write-Log "Processing $($Operations.Count) operations individually as fallback" -Level INFO
    
    foreach ($operation in $Operations) {
        try {
            switch ($operation.Operation) {
                "UpdateContentType" {
                    Set-PnPListItem -List $operation.Data.ListName -Identity $operation.Item.Id -ContentType $operation.Data.ContentTypeName -UpdateType "SystemUpdate" -Force
                    Write-Log "Content type updated (individual): $($operation.Item.FieldValues.FileRef)" -Level SUCCESS
                }
                "UpdateMetadata" {
                    Set-PnPListItem -List $operation.Data.ListName -Identity $operation.Item.Id -Values @{$operation.Data.ColumnName = $operation.Data.TermIdList } -UpdateType "SystemUpdate" -Force
                    Write-Log "Managed metadata updated (individual): $($operation.Data.FolderPath)" -Level SUCCESS
                }
            }
            $script:SuccessCount++
        }
        catch {
            Write-Log "Error processing individual operation for $($operation.Item.FieldValues.FileRef): $($_.Exception.Message)" -Level ERROR
            $script:ErrorCount++
        }
        finally {
            $script:ProcessedCount++
        }
    }
}

function Process-CsvData {
    param([array]$CsvData)
    
    Write-Log "Processing $($CsvData.Count) CSV entries" -Level INFO
    
    $batchItems = @()
    $currentBatch = 1
    
    foreach ($csvEntry in $CsvData) {
        try {
            $listName = $csvEntry.Library
            $folderPath = $csvEntry.FolderPath
            $columnName = $csvEntry."Sharepoint Column Name"
            $terms = $csvEntry.Terms
            $termGroup = $csvEntry."Term Group"
            $termSet = $csvEntry."Term Set"
            $managedMetadataValues = ($terms -split ";").Trim()
            
            # Validate required data
            if ([string]::IsNullOrEmpty($listName) -or [string]::IsNullOrEmpty($folderPath)) {
                Write-Log "Skipping entry with missing required data" -Level WARN
                $script:SkippedCount++
                continue
            }
            
            # Check if content type exists in list
            $contentType = Get-CachedContentType -ListName $listName -ContentTypeName $TargetContentTypeName
            if (-not $contentType) {
                Write-Log "Content Type '$TargetContentTypeName' not found in list '$listName' - skipping" -Level ERROR
                $script:ErrorCount++
                continue
            }
            
            # Get terms
            $termIdList = Get-CachedTerms -ManagedMetadataValues $managedMetadataValues -TermSet $termSet -TermGroup $termGroup
            if ($termIdList.Count -eq 0) {
                Write-Log "No valid terms found for folder '$folderPath' - skipping metadata update" -Level WARN
                $script:SkippedCount++
                continue
            }
            
            # Get folder items
            $folderResult = Get-FolderItems -ListName $listName -FolderPath $folderPath
            if (-not $folderResult) {
                $script:ErrorCount++
                continue
            }
            
            # Process each folder item
            foreach ($folderItem in $folderResult.Items) {
                # Check if content type update is needed
                $itemWithContentType = Get-PnPListItem -List $listName -Id $folderItem.Id -IncludeContentType
                
                if ($itemWithContentType.ContentType.Name -ne $TargetContentTypeName) {
                    $batchItems += @{
                        Item      = $folderItem
                        Operation = "UpdateContentType"
                        Data      = @{
                            ListName        = $listName
                            ContentTypeName = $TargetContentTypeName
                        }
                    }
                }
                else {
                    Write-Log "Item '$($folderItem.FieldValues.FileRef)' already has target content type - skipping" -Level DEBUG
                    $script:SkippedCount++
                }
                
                # Always add metadata update (for the main folder item)
                if ($folderItem -eq $folderResult.Items[0]) {
                    # Only for the first item (main folder)
                    $batchItems += @{
                        Item      = $folderItem
                        Operation = "UpdateMetadata"
                        Data      = @{
                            ListName   = $listName
                            ColumnName = $columnName
                            TermIdList = $termIdList
                            FolderPath = $folderResult.FullPath
                        }
                    }
                }
            }
            
            # Process batch when it reaches the batch size
            if ($batchItems.Count -ge $BatchSize) {
                Process-ItemsBatch -BatchItems $batchItems -BatchNumber $currentBatch
                $batchItems = @()
                $currentBatch++
            }
        }
        catch {
            Write-Log "Error processing CSV entry: $($_.Exception.Message)" -Level ERROR
            $script:ErrorCount++
        }
    }
    
    # Process remaining items in final batch
    if ($batchItems.Count -gt 0) {
        Process-ItemsBatch -BatchItems $batchItems -BatchNumber $currentBatch
    }
}

function Show-Progress {
    param([int]$Current, [int]$Total)
    
    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    Write-Progress -Activity "Processing Managed Metadata Updates" -Status "$Current of $Total items processed ($percent%)" -PercentComplete $percent
}

function Write-Summary {
    Write-Log "=== PROCESSING SUMMARY ===" -Level INFO
    Write-Log "Total Processed: $script:ProcessedCount" -Level INFO
    Write-Log "Successful Updates: $script:SuccessCount" -Level SUCCESS
    Write-Log "Skipped Items: $script:SkippedCount" -Level INFO
    Write-Log "Errors: $script:ErrorCount" -Level ERROR
    Write-Log "Log file: $script:LogFile" -Level INFO
    
    # Cache statistics
    Write-Log "=== CACHE STATISTICS ===" -Level INFO
    Write-Log "Cached Lists: $($script:ListCache.Keys.Count)" -Level INFO
    Write-Log "Cached Content Types: $($script:ContentTypeCache.Keys.Count)" -Level INFO
    Write-Log "Cached Terms: $($script:TermCache.Keys.Count)" -Level INFO
    
    if ($script:ErrorCount -eq 0) {
        Write-Log "All operations completed successfully!" -Level SUCCESS
    }
    elseif ($script:SuccessCount -gt 0) {
        Write-Log "Completed with some errors. Check log file for details." -Level WARN
    }
    else {
        Write-Log "No successful operations. Check configuration and log file." -Level ERROR
    }
}

# Main execution
try {
    Initialize-Script
    Connect-ToSite
    
    # Load and validate CSV data
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "Loaded $($csvData.Count) entries from CSV" -Level INFO
    
    # Validate CSV structure
    $requiredColumns = @('Library', 'FolderPath', 'Sharepoint Column Name', 'Terms', 'Term Group', 'Term Set')
    $csvColumns = $csvData[0].PSObject.Properties.Name
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
    
    if ($missingColumns.Count -gt 0) {
        throw "Missing required CSV columns: $($missingColumns -join ', ')"
    }
    
    # Process CSV data with batching
    Process-CsvData -CsvData $csvData
    
    Write-Progress -Activity "Processing Managed Metadata Updates" -Completed
    Write-Summary
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level ERROR
    Write-Summary
    exit 1
}
finally {
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnect errors
    }
}