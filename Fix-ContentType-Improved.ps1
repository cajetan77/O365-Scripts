<#
.SYNOPSIS
    Updates content type of list items based on CSV input using PnP batch operations for improved performance.

.DESCRIPTION
    This script processes a CSV file containing SiteURL, Directory, and FileName columns.
    It updates the content type of documents to "RBNZ Migrated Document" using batch operations
    for better performance and reliability.

.PARAMETER CertificateThumbprint
    The certificate thumbprint for authentication

.PARAMETER ApplicationId
    The application ID for authentication

.PARAMETER TenantId
    The tenant ID for authentication

.PARAMETER CsvPath
    Path to the CSV file containing items to process

.PARAMETER UpdatedContentType
    The content type to set on the items

.PARAMETER BatchSize
    Number of operations to batch together (default: 100)

.EXAMPLE
    .\Fix-ContentType-Improved.ps1 -CsvPath ".\SiteContentTypeFix.csv"
#>

param(
    [string]$CertificateThumbprint = "F4CC5E6DF6A53A34414AEE21EF66471450D4ECBE",
    [string]$ApplicationId = "054b783c-84cd-4f4c-8c92-ca35a6828679",
    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",   
    [string]$CsvPath = ".\SiteContentTypeFix.csv",
    [string]$UpdatedContentType = "RBNZ Migrated Document",
    [int]$BatchSize = 100
) 

# Global variables
$script:LogFile = "$PSScriptRoot\script_FixContentType_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:ProcessedCount = 0
$script:SuccessCount = 0
$script:ErrorCount = 0

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
    Write-Log "Starting Fix-ContentType script with batch operations" -Level INFO
    Write-Log "CSV Path: $CsvPath" -Level INFO
    Write-Log "Target Content Type: $UpdatedContentType" -Level INFO
    Write-Log "Batch Size: $BatchSize" -Level INFO
    Write-Log "Log File: $script:LogFile" -Level INFO
    
    # Validate CSV file exists
    if (-not (Test-Path $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level ERROR
        throw "CSV file not found"
    }
}

function Get-LibraryFromPath {
    param([string]$ServerRelativeUrl)
    
    $segments = $ServerRelativeUrl -split "/"
    if ($segments.Count -ge 4) {
        return $segments[3]
    }
    return $null
}

function Ensure-ContentTypeInLibrary {
    param(
        [string]$LibraryName,
        [string]$ContentTypeName,
        [object]$Batch
    )
    
    try {
        # Check if content type exists in library
        $existingContentType = Get-PnPContentType -List $LibraryName -Identity $ContentTypeName -ErrorAction SilentlyContinue
        
        if ($null -eq $existingContentType) {
            Write-Log "Adding content type '$ContentTypeName' to library '$LibraryName'" -Level INFO
            
            # Add content type to library (not batched as it's a prerequisite)
            Add-PnPContentTypeToList -List $LibraryName -ContentType $ContentTypeName -ErrorAction Stop
            
            # Small delay to ensure content type is available
            Start-Sleep -Seconds 2
            
            # Verify it was added
            $existingContentType = Get-PnPContentType -List $LibraryName -Identity $ContentTypeName -ErrorAction SilentlyContinue
            
            if ($null -eq $existingContentType) {
                throw "Failed to add content type '$ContentTypeName' to library '$LibraryName'"
            }
            
            Write-Log "Successfully added content type '$ContentTypeName' to library '$LibraryName'" -Level SUCCESS
        }
        
        return $existingContentType
    }
    catch {
        Write-Log "Error ensuring content type in library '$LibraryName': $($_.Exception.Message)" -Level ERROR
        throw
    }
}

function Process-SiteItems {
    param(
        [string]$SiteUrl,
        [array]$SiteItems
    )
    
    Write-Log "Processing $($SiteItems.Count) items for site: $SiteUrl" -Level INFO
    
    try {
        # Connect to site
        Connect-PnPOnline -Url $SiteUrl -ClientId $ApplicationId -Thumbprint $CertificateThumbprint -Tenant $TenantId
        $web = Get-PnPWeb | Select-Object Title
        Write-Log "Connected to site: $($web.Title)" -Level SUCCESS
        
        # Group items by library for batch processing
        $itemsByLibrary = $SiteItems | Group-Object { Get-LibraryFromPath $_.Directory }
        
        foreach ($libraryGroup in $itemsByLibrary) {
            $libraryName = $libraryGroup.Name
            $libraryItems = $libraryGroup.Group
            
            if ([string]::IsNullOrEmpty($libraryName)) {
                Write-Log "Invalid library path for items, skipping" -Level WARN
                continue
            }
            
            Write-Log "Processing $($libraryItems.Count) items in library: $libraryName" -Level INFO
            
            try {
                # Ensure content type exists in library
                $contentType = Ensure-ContentTypeInLibrary -LibraryName $libraryName -ContentTypeName $UpdatedContentType -Batch $null
                
                # Process items in batches
                $batches = @()
                for ($i = 0; $i -lt $libraryItems.Count; $i += $BatchSize) {
                    $batchItems = $libraryItems[$i..([Math]::Min($i + $BatchSize - 1, $libraryItems.Count - 1))]
                    $batches += ,$batchItems
                }
                
                foreach ($batchItems in $batches) {
                    Process-ItemBatch -LibraryName $libraryName -Items $batchItems -ContentType $contentType
                }
            }
            catch {
                Write-Log "Error processing library '$libraryName': $($_.Exception.Message)" -Level ERROR
                $script:ErrorCount += $libraryItems.Count
            }
        }
    }
    catch {
        Write-Log "Error connecting to site '$SiteUrl': $($_.Exception.Message)" -Level ERROR
        $script:ErrorCount += $SiteItems.Count
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore disconnect errors
        }
    }
}

function Process-ItemBatch {
    param(
        [string]$LibraryName,
        [array]$Items,
        [object]$ContentType
    )
    
    Write-Log "Processing batch of $($Items.Count) items in library: $LibraryName" -Level DEBUG
    
    # Create a new batch
    $batch = New-PnPBatch
    $batchOperations = @()
    
    foreach ($item in $Items) {
        try {
            $filePath = "$($item.Directory)/$($item.FileName)"
            
            # Find the list item
            $listItem = Get-PnPListItem -List $LibraryName -Includes "ContentType" -PageSize 500 | 
                Where-Object { $_.FieldValues.FileRef -eq $filePath }
            
            if ($null -eq $listItem) {
                Write-Log "Item not found: $filePath" -Level WARN
                $script:ErrorCount++
                continue
            }
            
            # Check if content type update is needed
            if ($listItem.ContentType.Name -eq $UpdatedContentType) {
                Write-Log "Item already has correct content type: $($item.FileName)" -Level DEBUG
                $script:ProcessedCount++
                continue
            }
            
            Write-Log "Current Content Type for $($item.FileName): $($listItem.ContentType.Name)" -Level DEBUG
            
            # Add to batch
            $batchOperations += @{
                ListItem = $listItem
                Item = $item
                FilePath = $filePath
            }
            
            # Add batch operation
            Set-PnPListItem -List $LibraryName -Identity $listItem.Id -ContentType $ContentType -UpdateType SystemUpdate -Batch $batch
            
        }
        catch {
            Write-Log "Error preparing item $($item.FileName) for batch: $($_.Exception.Message)" -Level ERROR
            $script:ErrorCount++
        }
    }
    
    # Execute batch if there are operations
    if ($batchOperations.Count -gt 0) {
        try {
            Write-Log "Executing batch with $($batchOperations.Count) operations" -Level INFO
            Invoke-PnPBatch -Batch $batch
            
            # Log success for each item in the batch
            foreach ($operation in $batchOperations) {
                Write-Log "Content Type updated for: $($operation.Item.FileName) in $($operation.Item.SiteUrl)" -Level SUCCESS
                $script:SuccessCount++
                $script:ProcessedCount++
            }
        }
        catch {
            Write-Log "Batch execution failed: $($_.Exception.Message)" -Level ERROR
            
            # Process items individually as fallback
            Write-Log "Falling back to individual processing for this batch" -Level WARN
            Process-ItemsIndividually -LibraryName $LibraryName -Operations $batchOperations -ContentType $ContentType
        }
    }
}

function Process-ItemsIndividually {
    param(
        [string]$LibraryName,
        [array]$Operations,
        [object]$ContentType
    )
    
    Write-Log "Processing $($Operations.Count) items individually as fallback" -Level INFO
    
    foreach ($operation in $Operations) {
        try {
            Set-PnPListItem -List $LibraryName -Identity $operation.ListItem.Id -ContentType $ContentType -UpdateType SystemUpdate
            Write-Log "Content Type updated (individual): $($operation.Item.FileName)" -Level SUCCESS
            $script:SuccessCount++
        }
        catch {
            Write-Log "Error updating content type (individual) for $($operation.Item.FileName): $($_.Exception.Message)" -Level ERROR
            $script:ErrorCount++
        }
        finally {
            $script:ProcessedCount++
        }
    }
}

function Show-Progress {
    param([int]$Current, [int]$Total)
    
    $percent = [math]::Round(($Current / $Total) * 100, 1)
    Write-Progress -Activity "Processing Content Type Updates" -Status "$Current of $Total items processed ($percent%)" -PercentComplete $percent
}

function Write-Summary {
    Write-Log "=== PROCESSING SUMMARY ===" -Level INFO
    Write-Log "Total Processed: $script:ProcessedCount" -Level INFO
    Write-Log "Successful Updates: $script:SuccessCount" -Level SUCCESS
    Write-Log "Errors: $script:ErrorCount" -Level ERROR
    Write-Log "Log file: $script:LogFile" -Level INFO
    
    if ($script:ErrorCount -eq 0) {
        Write-Log "All operations completed successfully!" -Level SUCCESS
    } elseif ($script:SuccessCount -gt 0) {
        Write-Log "Completed with some errors. Check log file for details." -Level WARN
    } else {
        Write-Log "No successful operations. Check configuration and log file." -Level ERROR
    }
}

# Main execution
try {
    Initialize-Script
    
    # Load and validate CSV data
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "Loaded $($csvData.Count) items from CSV" -Level INFO
    
    # Validate CSV structure
    $requiredColumns = @('SiteUrl', 'Directory', 'FileName')
    $csvColumns = $csvData[0].PSObject.Properties.Name
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
    
    if ($missingColumns.Count -gt 0) {
        throw "Missing required CSV columns: $($missingColumns -join ', ')"
    }
    
    # Group items by site for efficient processing
    $itemsBySite = $csvData | Group-Object SiteUrl
    Write-Log "Processing $($itemsBySite.Count) unique sites" -Level INFO
    
    $currentSite = 0
    foreach ($siteGroup in $itemsBySite) {
        $currentSite++
        Show-Progress -Current $currentSite -Total $itemsBySite.Count
        
        $siteUrl = $siteGroup.Name
        $siteItems = $siteGroup.Group
        
        Process-SiteItems -SiteUrl $siteUrl -SiteItems $siteItems
    }
    
    Write-Progress -Activity "Processing Content Type Updates" -Completed
    Write-Summary
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level ERROR
    Write-Summary
    exit 1
}