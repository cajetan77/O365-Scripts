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

Connect-PnPOnline -Url "https://caje77sharepoint.sharepoint.com/sites/App-4" -Interactive -ClientId "66a1852a-1f21-46a2-ad58-35fc4c3f1530"

$outlook = $null

try {
    $outlook = New-Object -ComObject Outlook.Application

    foreach ($msgFile in $msgFiles) {
        Write-Host "Uploading $($msgFile.Name)"

        $metadata = Get-MsgMetadata `
            -MsgFile $msgFile.FullName `
            -Outlook $outlook

        Add-PnPFile `
            -Path $msgFile.FullName `
            -Folder "ArchiveEmails" `
            -Values $metadata
    }
}
finally {
    if ($outlook) {
        Remove-ComObject -ComObject $outlook
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}