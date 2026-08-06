<#
.SYNOPSIS
    Adds or removes approved Users on Distribution Lists from the SharePoint Groups list.

.DESCRIPTION
    1. Read Groups list items where Approved = Yes and Distribution List is set
    2. Task = Add    -> Add-DistributionGroupMember
       Task = Remove -> Remove-DistributionGroupMember
    3. Clear Approved on success
    4. Save log to Documents\GroupAccessLogs as UserName-Group-Date.txt

.NOTES
    Variables: SHAREPOINT_SITE_URL, EXCHANGE_ORGANIZATION
    Optional: SHAREPOINT_FOLDER_PATH_GROUPACCESS (default Shared Documents/GroupAccessLogs)
    Modules: PnP.PowerShell, ExchangeOnlineManagement
    MI needs Exchange.ManageAsApp + Exchange Administrator (or equivalent)
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
$logGroup = 'DistributionList'
$logUploaded = $false

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

    $ex = $ErrorRecord.Exception
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($ex -and -not [string]::IsNullOrWhiteSpace($ex.Message)) { [void]$parts.Add([string]$ex.Message) }
    if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace($ErrorRecord.ErrorDetails.Message)) {
        [void]$parts.Add([string]$ErrorRecord.ErrorDetails.Message)
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.FullyQualifiedErrorId)) {
        [void]$parts.Add([string]$ErrorRecord.FullyQualifiedErrorId)
    }
    if ($ErrorRecord.CategoryInfo -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.CategoryInfo.Reason)) {
        [void]$parts.Add([string]$ErrorRecord.CategoryInfo.Reason)
    }
    if ($ex -and $ex.InnerException -and -not [string]::IsNullOrWhiteSpace($ex.InnerException.Message)) {
        [void]$parts.Add([string]$ex.InnerException.Message)
    }

    if ($parts.Count -eq 0) {
        $asText = [string]$ErrorRecord
        if (-not [string]::IsNullOrWhiteSpace($asText)) { return $asText.Trim() }
        return 'Unknown Exchange error (empty message)'
    }

    return ($parts | Select-Object -First 3) -join ' | '
}

function Save-RunLog {
    if ($script:logUploaded) { return }
    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) { return }

    try {
        Connect-PnPOnline -Url $script:siteUrl -ManagedIdentity
        $fileName = '{0}-{1}-{2}.txt' -f (Get-SafeName $script:logUser), (Get-SafeName $script:logGroup), $script:stamp
        $temp = Join-Path $env:TEMP $fileName
        ($script:log -join [Environment]::NewLine) | Set-Content -LiteralPath $temp -Encoding UTF8
        Resolve-PnPFolder -SiteRelativePath $script:folder | Out-Null
        Add-PnPFile -Path $temp -Folder $script:folder -NewFileName $fileName | Out-Null
        Remove-Item $temp -Force -ErrorAction SilentlyContinue
        $script:logUploaded = $true
        Write-Log "Log saved: $($script:folder)/$fileName"
    }
    catch {
        Write-Output ("[{0}] WARN: could not upload log: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $_.Exception.Message)
    }
}

Write-Log 'Starting Request-DistributionListAccess'
Write-Log "List='$ListTitle' | Approved=Yes | Distribution List only | Task=Add or Remove"

$toProcess = @()
$outcomes = @()

try {
    # --- 1) Read approved Distribution List requests ---
    Import-Module PnP.PowerShell -ErrorAction Stop
    Connect-PnPOnline -Url $siteUrl -ManagedIdentity

    foreach ($item in @(Get-PnPListItem -List $ListTitle -PageSize 2000)) {
        if ([string]$item['Approved'] -ne 'Yes') { continue }

        $dl = [string]$item['DistributionList']
        if (-not $dl) { $dl = [string]$item['Distribution List'] }
        if ([string]::IsNullOrWhiteSpace($dl)) { continue }

        $mesg = [string]$item['MailEnabled']
        if (-not $mesg) { $mesg = [string]$item['Mail Enabled'] }

        $user = $item['User']
        if ($user -is [array]) { $user = $user[0] }

        $userEmail = $null
        $userName = $null
        if ($user) {
            $userEmail = @($user.Email, $user.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1
            $userName = @($user.LookupValue, $user.Email, $userEmail) | Where-Object { $_ } | Select-Object -First 1
        }

        $task = [string]$item['Task']
        if ($task) { $task = $task.Trim() }

        $toProcess += [pscustomobject]@{
            Id          = [int]$item.Id
            UserEmail   = if ($userEmail) { $userEmail.Trim() } else { $null }
            UserName    = if ($userName) { $userName.Trim() } else { $userEmail }
            Group       = $dl.Trim()
            Task        = $task
            HasMesgAlso = -not [string]::IsNullOrWhiteSpace($mesg)
        }
    }

    Write-Log "Found $($toProcess.Count) approved Distribution List item(s)"

    if ($toProcess.Count -eq 0) {
        Write-Log 'Nothing to do'
    }
    else {
        if ($toProcess.Count -eq 1) {
            $logUser = $toProcess[0].UserName
            $logGroup = $toProcess[0].Group
        }

        # --- 2) Add / Remove on distribution lists ---
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Log "Connecting to Exchange ($org)"
        Connect-ExchangeOnline -ManagedIdentity -Organization $org -ShowBanner:$false

        foreach ($req in $toProcess) {
            $name = if ($req.UserName) { $req.UserName } else { 'Unknown user' }
            $task = $req.Task

            if ($req.HasMesgAlso) {
                Write-Log "'$name' Failed: set Distribution List only (clear Mail Enabled)"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }
            if (-not $req.UserEmail) {
                Write-Log "'$name' Failed: User email missing"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }
            if ($task -notin @('Add', 'Remove')) {
                Write-Log "'$name' Failed: Task must be Add or Remove (got '$task')"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }

            Write-Log "'$name' ($($req.UserEmail)) $task DL '$($req.Group)'"

            try {
                $exoGroup = Get-DistributionGroup -Identity $req.Group -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-Log "'$name' Failed: group '$($req.Group)' not found in Exchange. $(Get-ErrorText $_)"
                Write-Log "Tip: use the group's email or exact Exchange display name."
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }

            $type = [string]$exoGroup.RecipientTypeDetails
            if ($type -ne 'MailUniversalDistributionGroup') {
                Write-Log "'$name' Failed: '$($req.Group)' is '$type', not a Distribution List (use Mail Enabled runbook for security groups)"
                $outcomes += [pscustomobject]@{ Id = $req.Id; Name = $name; Ok = $false }
                continue
            }

            $errText = $null
            try {
                if ($task -eq 'Add') {
                    Add-DistributionGroupMember -Identity $exoGroup.Identity -Member $req.UserEmail -BypassSecurityGroupManagerCheck -ErrorAction Stop | Out-Null
                }
                else {
                    Remove-DistributionGroupMember -Identity $exoGroup.Identity -Member $req.UserEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop | Out-Null
                }
            }
            catch {
                $errText = Get-ErrorText $_
            }

            $alreadyAdd = $errText -and ($errText -match 'already exist|already a member|already present')
            $alreadyRemove = $errText -and ($errText -match "isn't a member|is not a member|couldn't find|could not find|not a member")

            if (-not $errText) {
                if ($task -eq 'Add') {
                    Write-Log "'$name' Added to DL '$($req.Group)'"
                }
                else {
                    Write-Log "'$name' Removed from DL '$($req.Group)'"
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

        foreach ($o in $outcomes) {
            if ($o.Ok) {
                Set-PnPListItem -List $ListTitle -Identity $o.Id -Values @{ Approved = $null } | Out-Null
                Write-Log "Cleared Approved for '$($o.Name)'"
            }
            else {
                Write-Log "Left Approved=Yes for '$($o.Name)' (retry later)"
            }
        }

        if ($outcomes.Count -eq 1 -and $outcomes[0].Ok) {
            $logUser = $outcomes[0].Name
            $logGroup = $outcomes[0].Group
        }

        Write-Log 'Uploading log file'
        Save-RunLog
    }

    $ok = @($outcomes | Where-Object Ok).Count
    $bad = @($outcomes | Where-Object { -not $_.Ok }).Count
    Write-Log "Done. success=$ok failed=$bad"
    if ($bad -gt 0) {
        # Surface a clear reason in AA output (still fails the job)
        throw "$bad request(s) failed — see Output stream / GroupAccessLogs"
    }
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    try { Save-RunLog } catch { }
    throw
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
    try { if (Get-PnPConnection -ErrorAction SilentlyContinue) { Disconnect-PnPOnline } } catch { }
}
