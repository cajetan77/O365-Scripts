<#
.SYNOPSIS
    Exports Entra user details: Department, Manager, Interactive and Non-Interactive last sign-in.

.DESCRIPTION
    Azure Automation runbook using system-assigned Managed Identity.
    Writes UserDetails.csv under $env:TEMP, then uploads (overwrites) to SharePoint.

.NOTES
    Permissions (Graph app roles on the MI):
      User.Read.All
      AuditLog.Read.All   (signInActivity — needs Entra ID P1/P2)
      Sites.ReadWrite.All (or SharePoint Sites.Selected / FullControl for PnP)

    Modules (same Graph version):
      Microsoft.Graph.Authentication
      Microsoft.Graph.Users
      PnP.PowerShell

    Automation Variables:
      SHAREPOINT_SITE_URL
      SHAREPOINT_FOLDER_PATH_USERDETAILS  (optional; default Shared Documents/UserDetails)

.EXAMPLE
    .\Get-UserDetails.ps1
#>
[CmdletBinding()]
Param()

$ErrorActionPreference = 'Stop'

$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
$SharePointFolderPath = Get-AutomationVariable -Name 'SHAREPOINT_FOLDER_PATH_USERDETAILS' -ErrorAction SilentlyContinue
if ([string]::IsNullOrWhiteSpace($SharePointFolderPath)) {
    $SharePointFolderPath = 'Shared Documents/UserDetails'
}
if ($SharePointFolderPath -notmatch '^(Shared Documents|Documents)(/|$)') {
    $SharePointFolderPath = "Shared Documents/$($SharePointFolderPath.Trim('/'))"
}

function Write-RunbookLog([string]$Message) {
    Write-Output ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Get-GraphProperty($Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    return $Object.$Name
}

Write-RunbookLog 'Starting Get-UserDetails'

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module PnP.PowerShell -ErrorAction Stop

Write-RunbookLog 'Connecting to Microsoft Graph (Managed Identity)...'
Connect-MgGraph -Identity -NoWelcome

$rows = [System.Collections.Generic.List[object]]::new()
# signInActivity needs AuditLog.Read.All + Entra P1/P2
$pageUri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,department,accountEnabled,signInActivity&$expand=manager($select=displayName,userPrincipalName)&$top=100'
$pageNumber = 0

while (-not [string]::IsNullOrWhiteSpace($pageUri)) {
    $pageNumber++
    Write-RunbookLog "Fetching users page $pageNumber..."
    $page = Invoke-MgGraphRequest -Method GET -Uri $pageUri

    foreach ($u in @((Get-GraphProperty $page 'value'))) {
        if ($null -eq $u) { continue }

        $manager = Get-GraphProperty $u 'manager'
        $signIn = Get-GraphProperty $u 'signInActivity'

        $rows.Add([PSCustomObject]@{
                DisplayName                      = [string](Get-GraphProperty $u 'displayName')
                UserPrincipalName                = [string](Get-GraphProperty $u 'userPrincipalName')
                Mail                             = [string](Get-GraphProperty $u 'mail')
                Department                       = [string](Get-GraphProperty $u 'department')
                AccountEnabled                   = [bool](Get-GraphProperty $u 'accountEnabled')
                Manager                          = [string](Get-GraphProperty $manager 'displayName')
                ManagerUPN                       = [string](Get-GraphProperty $manager 'userPrincipalName')
                LastInteractiveSignInDateTime    = Get-GraphProperty $signIn 'lastSignInDateTime'
                LastNonInteractiveSignInDateTime = Get-GraphProperty $signIn 'lastNonInteractiveSignInDateTime'
            })
    }

    $pageUri = [string](Get-GraphProperty $page '@odata.nextLink')
}

Write-RunbookLog "Retrieved $($rows.Count) user(s)"

$csvPath = Join-Path $env:TEMP 'UserDetails.csv'
@($rows.ToArray()) | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

Write-RunbookLog "Uploading to SharePoint: $SharePointSiteUrl / $SharePointFolderPath"
Connect-PnPOnline -Url $SharePointSiteUrl -ManagedIdentity
try {
    Add-PnPFile -Path $csvPath -Folder $SharePointFolderPath -NewFileName 'UserDetails.csv' | Out-Null
    Write-RunbookLog "Uploaded/updated: $SharePointFolderPath/UserDetails.csv"
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

Remove-Item -LiteralPath $csvPath -Force -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Write-RunbookLog "Done. users=$($rows.Count)"
