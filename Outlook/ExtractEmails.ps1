
$pstName = "caje77@keiratheapp.com"   # change if needed
$pstPath = "D:\Powershell\O365 Scripts\Outlook\backup.pst"
$exportPath = "D:\Powershell\O365 Scripts\Outlook"


$outlook = New-Object -ComObject Outlook.Application
#$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")

$pst = $ns.Folders.Item($pstName)
$inbox = $pst.Folders.Item("Inbox")

$count = $inbox.Items.Count
Write-Host "Found $count emails"

for ($i = 1; $i -le $count; $i++) {
    try {
        $mail = $inbox.Items.Item($i)

        if ($mail.MessageClass -like "IPM.Note*") {

            $subject = if ([string]::IsNullOrWhiteSpace($mail.Subject)) {
                "NoSubject"
            } else {
                $mail.Subject
            }

            $safeName = $subject -replace '[\\/:*?"<>|]', "_"

            if ($safeName.Length -gt 80) {
                $safeName = $safeName.Substring(0,80)
            }

            $fileName = "{0:yyyyMMdd_HHmmss}_{1}.msg" -f $mail.ReceivedTime, $safeName
            $filePath = Join-Path $exportPath $fileName

            $mail.SaveAs($filePath, 3)

            Write-Host "Exported: $fileName"
        }
    }
    catch {
        Write-Warning "Failed item $i"
    }
}