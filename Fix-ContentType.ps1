# The script will update content type of list items based on CSV input
# CSV should have columns: SiteURL, Directory, FileName 
#It will replace content type  of the Documents present  with "RBNZ Migrated Document", if it exists in the library, else it will add the content type to the library and then update the item

param(

    [string]$CertificateThumbprint = "F4CC5E6DF6A53A34414AEE21EF66471450D4ECBE",
    [string]$ApplicationId = "054b783c-84cd-4f4c-8c92-ca35a6828679",
    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",   
    [string]$csv = ".\SiteContentTypeFix.csv",
    [string]$updatedContentType = "RBNZ Migrated Document"  
   
) 

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO",

        [string]$LogFile = "$PSScriptRoot\script_FixContentType.log"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"

    # Write to console
    Write-Host $logEntry

    # Append to log file
    Add-Content -Path $LogFile -Value $logEntry
}


foreach ($site in (Import-Csv -Path $csv)) {
   
    try {
        Connect-PnPOnline -Url $site.SiteUrl -ClientId $ApplicationId -Thumbprint $CertificateThumbprint -Tenant $TenantId
        $web = Get-PnPWeb | Select Title    
    }
    catch {
        Write-Log -Message ("Error connecting to site: " + $site.SiteUrl) -Level ERROR
        continue

    }

    $serverRelativeUrl= $site.Directory

    

$segments = $serverRelativeUrl -split "/"

# Get the segment between the 3rd and 4th slash (index 3)
$library = $segments[3]
$filePath =  "$($site.Directory)/$($site.FileName)"

$listItem = Get-PnPListItem -List $library -Includes "ContentType" -PageSize 500 | Where-Object { $_.FieldValues.FileRef -eq $filePath }


    if($listItem -ne $null) {

        Write-Log -Message ("Current Content Type: "  + $listItem.ContentType.Name) -Level INFO
        $migcont=Get-PnPContentType -List $library -Identity $updatedContentType -ErrorAction SilentlyContinue
        if($migcont -ne $null) {
            try {

            Set-PnPListItem -List $library -Identity $listItem.Id -ContentType $migcont -UpdateType SystemUpdate
            Write-Log -Message ("Content Type for item " + $site.FileName + " in site " + $site.SiteUrl + " updated to 'Migrated Document'") -Level INFO
    
            }
            catch {
                Write-Log -Message ("Error updating content type for item " + $site.FileName + " in site " + $site.SiteUrl) -Level ERROR
                continue
            }
                    }
        else {
             Write-Log -Message ("Content Type 'Migrated Document' not found in library " + $library + " in site " + $site.SiteUrl) -Level WARN
            
            Add-PnPContentTypeToList -List $library -ContentType $updatedContentType -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $migcont=Get-PnPContentType -List $library -Identity $updatedContentType -ErrorAction SilentlyContinue
            
            if($migcont -ne $null) {
                try {

                Set-PnPListItem -List $library -Identity $listItem.Id -ContentType $migcont -UpdateType SystemUpdate
                Write-Log -Message ("Content Type for item " + $site.FileName + " in site " + $site.SiteUrl + " updated to 'Migrated Document'") -Level INFO
        
                }
                catch {
                    Write-Log -Message ("Error updating content type for item " + $site.FileName + " in site " + $site.SiteUrl) -Level ERROR
                    continue
                }
            }
            else {
                 Write-Log -Message ("Content Type 'RBNZ Migrated Document' still not found in library " + $library + " in site " + $site.SiteUrl) -Level WARN
            }
            
           
        }

    }
    else {
        Write-Log -Message ("Item not found at " + $site.Directory) -Level WARN
    }

}
