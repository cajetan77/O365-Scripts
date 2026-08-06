<#
.SYNOPSIS
    Adds or removes approved Users from SharePoint Groups list on DL / mail-enabled groups.

.DESCRIPTION
    1. Read Groups list items where Approved = Yes
    2. Task = Add  -> add User to Distribution List or Mail Enabled (one only)
       Task = Remove -> remove User from that group
    3. Clear Approved on success
    4. Save log to Documents\GroupAccessLogs as UserName-Group-Date.txt

.NOTES
    Variables: SHAREPOINT_SITE_URL, EXCHANGE_ORGANIZATION
    Optional: SHAREPOINT_FOLDER_PATH_GROUPACCESS (default Shared Documents/GroupAccessLogs)
    Modules: PnP.PowerShell, ExchangeOnlineManagement
#>
[CmdletBinding()]
Param(
    [string]$ListTitle = 'Groups'
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'Continue'

$log = [System.Collections.Generic.List[string]]::new()
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logUser = 'Batch'
$logGroup = 'Approved'

$siteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
$org = Get-AutomationVariable -Name 'EXCHANGE_ORGANIZATION' -ErrorAction Stop
$folder = Get-AutomationVariable -Name 'SHAREPOINT_FOLDER_PATH_GROUPACCESS' -ErrorAction SilentlyContinue
if ([string]::IsNullOrWhiteSpace($folder)) { $folder = 'Shared Documents/GroupAccessLogs' }
if ($folder -notmatch '^(Shared Documents|Documents)(/|$)') { $folder = "Shared Documents/$($folder.Trim('/'))" }

function Write-Log([string]$Message) {
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $script:log.Add($line)
    Write-Output $line
}

function Get-SafeName([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'Unknown' }
    $t = $Value.Trim()
    if ($t -match '@') { $t = ($t -split '@')[0] }
    $t = $t -replace '[\\/:*?"<>|#%&{}@$+|=!'']', '-' -replace '\s+', '-'
    if ($t.Length -gt 60) { $t = $t.Substring(0, 60) }
    return $t.Trim('-')
}

function Get-ErrorText($ErrorRecord) {
    if (-not $ErrorRecord) { return 'Unknown Exchange error (no error details returned)' }

    $parts = @(
        [string]$ErrorRecord.Exception.Message
        [string]$ErrorRecord.ErrorDetails.Message
        [string]$ErrorRecord.FullyQualifiedErrorId
        [string]$ErrorRecord.CategoryInfo.Reason
        [string]$ErrorRecord.Exception.InnerException.Message
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($parts.Count -eq 0) {
        $asText = [string]$ErrorRecord
        if (-not [string]::IsNullOrWhiteSpace($asText)) { return $asText.Trim() }
        return 'Unknown Exchange error (empty message)'
    }

    return ($parts | Select-Object -First 3) -join ' | '
}

#Write-Log 'Starting Request-GroupAccess'
#Write-Log "List='$ListTitle' | Approved=Yes only | Task=Add or Remove"

$toProcess = @()
$outcomes = @()

try {
    # --- 1) Read approved requests from SharePoint ---
    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline -Url $siteUrl -ManagedIdentity

    foreach ($item in @(Get-PnPListItem -List $ListTitle -PageSize 2000)) {
        if ([string]$item['Approved'] -ne 'Yes') { continue }

        $user = $item['User']
        if ($user -is [array]) { $user = $user[0] }

        $userEmail = $null
        $userName = $null
        if ($user) {
            $userEmail = @($user.Email, $user.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1
            $userName = @($user.LookupValue, $user.Email, $userEmail) | Where-Object { $_ } | Select-Object -First 1
        }

        $dl = [string]$item['DistributionList']
        if (-not $dl) { $dl = [string]$item['Distribution List'] }
        $mesg = [string]$item['MailEnabled']
        if (-not $mesg) { $mesg = [string]$item['Mail Enabled'] }

        $group = $null
        if ($dl -and -not $mesg) { $group = $dl.Trim() }
        elseif ($mesg -and -not $dl) { $group = $mesg.Trim() }

        $task = [string]$item['Task']
        if ($task) { $task = $task.Trim() }

        $toProcess += [pscustomobject]@{
            Id        = [int]$item.Id
            UserEmail = if ($userEmail) { $userEmail.Trim() } else { $null }
            UserName  = if ($userName) { $userName.Trim() } else { $userEmail }
            Group     = $group
            Task      = $task
        }
    }

    Write-Log "Found $($toProcess.Count) approved item(s)"

    if ($toProcess.Count -eq 0) {
        Write-Log 'Nothing to do'
    }
    else {
        if ($toProcess.Count -eq 1) {
            $logUser = $toProcess[0].UserName
            $logGroup = $toProcess[0].Group
        }

        # --- 2) Add / Remove in Exchange ---
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Log "Connecting to Exchange ($org)"
        Connect-ExchangeOnline -ManagedIdentity -Organization $org -ShowBanner:$false

        foreach ($req in $toProcess) {
            $name = if ($req.UserName) { $req.UserName } else { 'Unknown user' }
            $task = $req.Task

            if (-not $req.UserEmail) {
                Write-Log "'$name' Failed: User email missing"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }
            if (-not $req.Group) {
                Write-Log "'$name' Failed: pick one of Distribution List or Mail Enabled"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }
            if ($task -notin @('Add', 'Remove')) {
                Write-Log "'$name' Failed: Task must be Add or Remove (got '$task')"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }

            Write-Log "'$name' ($($req.UserEmail)) $task '$($req.Group)'"

            $exoGroup = $null
            $lookupErr = $null
            $exoGroup = Get-DistributionGroup -Identity $req.Group -ErrorAction SilentlyContinue -ErrorVariable lookupErr
            if (-not $exoGroup) {
                $reason = Get-ErrorText ($lookupErr | Select-Object -First 1)
                Write-Log "'$name' Failed: group '$($req.Group)' not found in Exchange. $reason"
                Write-Log "Tip: use the group's email address or exact Exchange display name in the SharePoint choice."
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }

            $err = $null
            if ($task -eq 'Add') {
                Add-DistributionGroupMember -Identity $exoGroup.Identity -Member $req.UserEmail -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue -ErrorVariable err | Out-Null
            }
            else {
                Remove-DistributionGroupMember -Identity $exoGroup.Identity -Member $req.UserEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable err | Out-Null
            }

            $errText = Get-ErrorText ($err | Select-Object -First 1)
            $alreadyAdd = $errText -match 'already exist|already a member|already present'
            $alreadyRemove = $errText -match "isn't a member|is not a member|couldn't find|could not find|not a member"

            if ((-not $err) -and $?) {
                if ($task -eq 'Add') {
                    Write-Log "'$name' Added to '$($req.Group)'"
                }
                else {
                    Write-Log "'$name' Removed from '$($req.Group)'"
                }
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $true; UserEmail = $req.UserEmail; Group = $req.Group; Task = $task }
            }
            elseif (($task -eq 'Add' -and $alreadyAdd) -or ($task -eq 'Remove' -and $alreadyRemove)) {
                if ($task -eq 'Add') {
                    Write-Log "'$name' Already a member of '$($req.Group)'"
                }
                else {
                    Write-Log "'$name' Already not a member of '$($req.Group)'"
                }
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $true; UserEmail = $req.UserEmail; Group = $req.Group; Task = $task }
            }
            else {
                Write-Log "'$name' Failed ($task): $errText"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
            }
        }

        try { Disconnect-ExchangeOnline -Confirm:$false } catch { }

        # --- 3) Clear Approved + upload log ---
        Connect-PnPOnline -Url $siteUrl -ManagedIdentity


        Write-Log 'Uploading log file'
        $fileName = '{0}-{1}-{2}.txt' -f (Get-SafeName $logUser), (Get-SafeName $logGroup), $stamp
        $temp = Join-Path $env:TEMP $fileName
        ($log -join [Environment]::NewLine) | Set-Content -LiteralPath $temp -Encoding UTF8
        Resolve-PnPFolder -SiteRelativePath $folder | Out-Null
        Add-PnPFile -Path $temp -Folder $folder -NewFileName $fileName | Out-Null
        Remove-Item $temp -Force -ErrorAction SilentlyContinue
        Write-Log "Log saved: $folder/$fileName"
    }

    $ok = @($outcomes | Where-Object Ok).Count
    $bad = @($outcomes | Where-Object { -not $_.Ok }).Count
    Write-Log "Done. success=$ok failed=$bad"
    if ($bad -gt 0) { throw "$bad request(s) failed — see log" }
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
    try { if (Get-PnPConnection -ErrorAction SilentlyContinue) { Disconnect-PnPOnline } } catch { }
}
