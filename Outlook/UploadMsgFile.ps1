function Remove-ComObject {
    param(
        [object]$ComObject
    )

    if ($ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Clear-OutlookComObjects {
    if (Get-Variable -Name mail -ErrorAction SilentlyContinue) {
        try {
            $mail.Close(0)
            Remove-ComObject -ComObject $mail
        }
        catch {
            Write-Verbose "No previous mail object needed to be released."
        }
    }

    if (Get-Variable -Name outlook -ErrorAction SilentlyContinue) {
        try {
            Remove-ComObject -ComObject $outlook
        }
        catch {
            Write-Verbose "No previous Outlook object needed to be released."
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-MsgMetadata {
    param(
        [Parameter(Mandatory)]
        [string]$MsgFile,

        [Parameter(Mandatory)]
        [object]$Outlook
    )

    $mail = $null

    try {
        $mail = $Outlook.Session.OpenSharedItem($MsgFile)

        return @{
            Title        = $mail.Subject
            Sender       = $mail.SenderEmailAddress
            Reciever     = $mail.To
            DateRecieved = $mail.ReceivedTime
        }
    }
    finally {
        if ($mail) {
            $mail.Close(0)
            Remove-ComObject -ComObject $mail
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Clear-OutlookComObjects

$msgFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.msg" -File

if (-not $msgFiles) {
    throw "No .msg files were found in '$PSScriptRoot'."
}

$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/App-4"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId "66a1852a-1f21-46a2-ad58-35fc4c3f1530"

$outlook = $null
$LibraryName = "ArchiveEmails"
try {
    $outlook = New-Object -ComObject Outlook.Application

    #foreach ($msgFile in $msgFiles) {
    #   Write-Host "Uploading $($msgFile.Name)"

    # $metadata = Get-MsgMetadata `
    #    -MsgFile $msgFile.FullName `
    #  -Outlook $outlook

    # Add-PnPFile `
    # -Path $msgFile.FullName `
    #  -Folder "ArchiveEmails" `
    #  -Values $metadata
    #}#
    $FolderRelativeUrl = "/sites/" + $SiteUrl.Split('/sites/')[1] + "/$LibraryName"

    $Query = @"
    <View Scope="FilesOnly">
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name="FileDirRef" />
                    <Value Type="Text">$FolderRelativeUrl</Value>
                </Eq>
            </Where>
        </Query>
    </View>
"@

    Write-Host "Fetching all files from '$LibraryName'..." -ForegroundColor Cyan
    $ListItems = Get-PnPListItem -List $LibraryName -Query $Query -PageSize 500
    foreach ($item in $ListItems) {
     
        if ($item.FieldValues.FSObjType -eq 0) {
            Write-Progress -Activity "Processing files" -Status "Processing file: $($item.FieldValues.FileLeafRef)" -PercentComplete ($item.Index / $ListItems.Count * 100)
            $FileRef = $Item.FieldValues.FileRef      # Complete relative URL of the file
            $FileName = $Item.FieldValues.FileLeafRef  # Just the file name
            $CreatedDate = [datetime]$Item.FieldValues.DateRecieved

            $YearStr = $CreatedDate.ToString("yyyy")
            $MonthStr = $CreatedDate.ToString("MM")
            $MonthStr = [DateTime]::ParseExact($MonthStr, "MM", $null).ToString("MMMM")
            
            # Build the exact target folder structure relative to the library root
            $TargetFolderPath = "$LibraryName/$YearStr/$MonthStr"
            $TargetFileRef = "/sites/" + $SiteUrl.Split('/sites/')[1] + "/$TargetFolderPath/$FileName"
            if ($FileRef -like "*/$YearStr/$MonthStr/$FileName") {
                Write-Host "Skipping '$FileName' - already in correct folder." -ForegroundColor Yellow
                continue
            }
            try {
                Write-Host "Processing file: $FileName (Created: $($CreatedDate.ToString('yyyy-MM-dd')))" -ForegroundColor White
                
                # 4. Dynamically provision the missing Year/Month folders if needed
                # Resolve-PnPFolder will return the folder or generate it safely if missing
                $Folder = Resolve-PnPFolder -SiteRelativePath $TargetFolderPath

                $TargetFileRef = "/sites/" + $SiteUrl.Split('/sites/')[1] + "/$TargetFolderPath"
                
                # 5. Relocate the file into the target directory
                Move-PnPFile -SourceUrl $FileRef -TargetUrl "$TargetFileRef" -Force
                Write-Host "Successfully moved '$FileName' to '$TargetFolderPath'" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to process file '$FileName'. Error: $_"
            }
        }
   
    }
}

finally {
    if ($outlook) {
        Remove-ComObject -ComObject $outlook
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}