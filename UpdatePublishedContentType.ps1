<#
.SYNOPSIS
    Updates published content types from Content Type Hub to target sites.
.DESCRIPTION
    Simple script to sync content types from Content Type Hub to target sites.
.PARAMETER ContentTypes
    Array of Content Type Names to update. Default: @("Koru Folder", "Koru Document")
.PARAMETER SitesCsvPath
    Path to CSV file containing site URLs. Default: .\SiteInventory.csv
.PARAMETER LogPath
    Path to the output log file. Default: .\ContentTypeUpdateLog.csv
.EXAMPLE
    .\UpdatePublishedContentType.ps1
    Updates default content types across all sites in SiteInventory.csv
.EXAMPLE
    .\UpdatePublishedContentType.ps1 -ContentTypes @("Document", "Folder")
    Updates specific content types from hub
#>

param(
    [string[]]$ContentTypes = @("Koru Folder", "KoruDocument"),
    [string]$SitesCsvPath = ".\SiteInventory.csv",
    [string]$LogPath = ".\ContentTypeUpdateLog.csv",
   
    [string]$TenantName = "caje77sharepoint",
    [string]$AdminUrl = "https://$TenantName-admin.sharepoint.com"
)

# Load configuration
$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.SharePointReportingAppId
$Thumbprint = $config.ThumbPrint

$ExcludedLists = @("Form Templates", "Master Page Gallery", "Style Library", "Site Assets")
$logEntries = New-Object System.Collections.Generic.List[object]

function Write-LogEntry {
    param(
        [string]$SiteUrl,
        [string]$Operation,
        [string]$ContentType = "",
        [string]$Library = "",
        [string]$Status,
        [string]$Message = ""
    )
    
    $logEntry = [pscustomobject]@{
        Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SiteUrl     = $SiteUrl
        Operation   = $Operation
        ContentType = $ContentType
        Library     = $Library
        Status      = $Status
        Message     = $Message
    }
    
    $logEntries.Add($logEntry)
}

Write-Host "Starting Content Type Update..." -ForegroundColor Green
Write-Host "Content Types: $($ContentTypes -join ', ')" -ForegroundColor Cyan
Write-Host ""

function Get-ContentTypeHub {
    param(
        [string[]]$ContentTypeNames,
        [string]$SiteUrl
    )
    
    try {
        # Get Content Type Hub URL
        Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
        #$ctHub = Get-PnPContentType
        #Disconnect-PnPOnline
        $contentTypesArray = $contentTypeNames.Split(",") | ForEach-Object { $_.Trim() }  
        $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
        Write-Host "Content Type Hub URL: $contentTypeHubUrl" -ForegroundColor Green
        $ctconnection = Connect-PnPOnline -Url $contentTypeHubUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
        $ctHub = Get-PnPContentType  -Connection $ctconnection

        Disconnect-PnPOnline

        $siteconnection = Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint
        # Enable feature for private channel sites
        #Enable-PnPFeature -Identity "73ef14b1-13a9-416b-a9b5-ececa2b0604c" -Scope Site -Force -ErrorAction SilentlyContinue
        
        $addedCount = 0
        foreach ($cts in $ctHub) {
            if ($contentTypesArray -contains $cts.Name) {
                try {
                    Add-PnPContentTypesFromContentTypeHub -ContentTypes $cts.Id.StringValue -Site $siteUrl -Connection $siteconnection
                    Write-Host "    ✓ Added: $($cts.Name)" -ForegroundColor Green
                    Write-LogEntry -SiteUrl $siteUrl -Operation "Add Content Type" -ContentType $cts.Name -Status "Success" -Message "Added from Content Type Hub"
                    $addedCount++
                }
                catch {
                    Write-Host "    ✗ Failed: $($cts.Name) - $($_.Exception.Message)" -ForegroundColor Red
                    Write-LogEntry -SiteUrl $siteUrl -Operation "Add Content Type" -ContentType $cts.Name -Status "Failed" -Message $_.Exception.Message
                }
            }
        }
        
        Disconnect-PnPOnline
        return $addedCount
    }
    catch {
        Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
        return 0
    }
}



# Read sites from CSV
if (-not (Test-Path $SitesCsvPath)) {
    Write-Host "ERROR: CSV file not found: $SitesCsvPath" -ForegroundColor Red
    exit 1
}

$sitesFromCsv = Import-Csv -Path $SitesCsvPath -Encoding UTF8
$sitesToProcess = @()

foreach ($siteRow in $sitesFromCsv) {
    $siteUrl = ""
    if (-not [string]::IsNullOrEmpty($siteRow.SiteUrl)) {
        $siteUrl = $siteRow.SiteUrl
    }
    elseif (-not [string]::IsNullOrEmpty($siteRow.'Site Collection Url')) {
        $siteUrl = $siteRow.'Site Collection Url'
    }
    
    if (-not [string]::IsNullOrEmpty($siteUrl)) {
        $sitesToProcess += $siteUrl.Trim()
    }
}

Write-Host "Processing $($sitesToProcess.Count) sites..." -ForegroundColor Cyan
Write-Host ""

# Process each site
$totalSites = $sitesToProcess.Count
$currentIndex = 0
$successfulSites = 0

foreach ($siteUrl in $sitesToProcess) {
    $currentIndex++
    Write-Host "[$currentIndex/$totalSites] Processing: $siteUrl" -ForegroundColor Cyan
    
    try {
        Write-LogEntry -SiteUrl $siteUrl -Operation "Site Processing" -Status "Started" -Message "Starting content type processing"
        
        # Add content types from hub
        $addedCount = Get-ContentTypeHub -ContentTypeNames $ContentTypes -SiteUrl $siteUrl
        
        if ($addedCount -gt 0) {
            $successfulSites++
            Write-LogEntry -SiteUrl $siteUrl -Operation "Site Processing" -Status "Success" -Message "$addedCount content types added from hub"
            
            
        }
        else {
            Write-LogEntry -SiteUrl $siteUrl -Operation "Site Processing" -Status "No Changes" -Message "No content types were added"
        }
        
        Write-Host "  Completed: $addedCount content types added" -ForegroundColor Green
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-LogEntry -SiteUrl $siteUrl -Operation "Site Processing" -Status "Error" -Message $_.Exception.Message
    }
    
    Write-Host ""
    
    # Small delay to avoid throttling
    if ($currentIndex -lt $totalSites) {
        Start-Sleep -Seconds 2
    }
}

# Export log to CSV
if ($logEntries.Count -gt 0) {
    try {
        # Ensure output directory exists
        $outputDirectory = Split-Path -Path $LogPath -Parent
        if ($outputDirectory -and -not (Test-Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        
        $logEntries | Sort-Object Timestamp | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        Write-Host "Log exported to: $LogPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Warning: Could not export log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Summary
Write-Host "========================================" -ForegroundColor Green
Write-Host "Content Type Update Completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Sites processed: $totalSites" -ForegroundColor Cyan
Write-Host "Sites with successful updates: $successfulSites" -ForegroundColor Cyan
Write-Host "Content types updated: $($ContentTypes -join ', ')" -ForegroundColor Cyan

# Show log summary
if ($logEntries.Count -gt 0) {
    $successOperations = ($logEntries | Where-Object { $_.Status -eq "Success" }).Count
    $failedOperations = ($logEntries | Where-Object { $_.Status -eq "Failed" }).Count
    
    Write-Host "Log entries created: $($logEntries.Count)" -ForegroundColor Cyan
    Write-Host "  - Successful operations: $successOperations" -ForegroundColor Green
    Write-Host "  - Failed operations: $failedOperations" -ForegroundColor Red
}

Write-Host ""
Write-Host "Script completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green