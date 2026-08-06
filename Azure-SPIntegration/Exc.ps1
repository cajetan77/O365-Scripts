Connect-ExchangeOnline
Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome

$dm = Get-DistributionGroupMember -Identity 'Z_Infra' |
    Where-Object { $_.ExternalDirectoryObjectId } |
    Select-Object -ExpandProperty ExternalDirectoryObjectId -Unique

foreach ($id in $dm) {
    try {
        $user = Get-MgUser -UserId $id -Property Id, DisplayName, UserPrincipalName, Mail -ErrorAction Stop
        [PSCustomObject]@{
            ExternalDirectoryObjectId = $user.Id
            DisplayName               = $user.DisplayName
            UserPrincipalName         = $user.UserPrincipalName
            Mail                      = $user.Mail
        }
    }
    catch {
        Write-Warning "No Graph user for ExternalDirectoryObjectId '$id': $($_.Exception.Message)"
    }
}
