<#
.SYNOPSIS
    Exports a complete Microsoft 365 licensing report to CSV.

.DESCRIPTION
    Produces four CSV files:
      - TenantSummary  : purchased, consumed, and available units per SKU
      - GroupLicenses  : licenses assigned to each Entra group (group-based licensing)
      - Users          : one row per user with identity, sign-in, and license summary
      - LicenseDetail  : one row per user/license with assigning group and service plans

    Requires an Entra app registration using client-secret/certificate (application) authentication.
    Grant admin consent for the application permissions listed in .NOTES.

.NOTES
    Microsoft Graph APPLICATION permissions (admin consent required):

    Permission              | Used by
    ------------------------|--------------------------------------------------
    User.Read.All           | Users, managers, sign-in activity, license detail
    Group.Read.All          | Groups with assigned licenses (group licensing)
    GroupMember.Read.All    | Member counts on licensing groups
    AuditLog.Read.All       | Sign-in activity (interactive / non-interactive)
    Organization.Read.All   | Tenant SKU summary (Get-MgSubscribedSku)

    Entra portal: App registrations > your app > API permissions > Add permission >
    Microsoft Graph > Application permissions > add each permission above > Grant admin consent.

    Sign-in activity (interactive / non-interactive) requires Entra ID P1 or P2 in the
    tenant for data to be populated; no additional Graph permission beyond User.Read.All.

    Large tenants (10k+ users): license names are resolved from the tenant SKU catalog
    (no per-user license API calls by default). Use -IncludeServicePlanDetail for disabled
    service plan columns (slower). Managers are loaded via Graph $batch requests.
#>
[CmdletBinding()]
Param(
    [ValidateSet('Member', 'Guest', 'All')]
    [string]$UserType = 'All',

    [switch]$LicensedUsersOnly,
    [switch]$UnlicensedUsersOnly,
    [switch]$EnabledUsersOnly,
    [switch]$DisabledUsersOnly,

    [switch]$IncludeServicePlanDetail,
    [switch]$SkipManager,
    [switch]$SkipGroupMemberCount,

    [int]$BatchSize = 20,
    [int]$BatchDelayMs = 200,

    [string]$TenantId = "764b46e8-d798-4ed3-87db-ae55ed7b0432",
    [string]$ClientId = "dc223b11-5ab5-4a33-988a-3474b25eb9be",
    [string]$ClientSecret = "",
    [string]$Thumbprint = '800DB610ED947E9251A199ABFEA40AED1738128E',
    [string]$ExportPath = ".\M365-LicensingReport",
    [string]$DomainFilter = 'cajesharepoint.onmicrosoft.com'
)

$ErrorActionPreference = 'Stop'

function Connect-MgGraphApp {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    Write-Host "Connecting to Microsoft Graph..."
    if ($Thumbprint) {
        # $credential = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $Thumbprint
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    }
    else {
        $secureSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
        $credential = [PSCredential]::new($ClientId, $secureSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
    }
    Set-MgRequestContext -Retries 5 -RetryDelay 10 | Out-Null
}


function Invoke-MgGraphBatch {
    param(
        [Parameter(Mandatory)]
        [array]$Requests,
        [int]$MaxRetries = 6
    )

    $body = @{ requests = $Requests } | ConvertTo-Json -Depth 6

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $result = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body $body
            return @($result.responses)
        }
        catch {
            $status = $null
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Min(60, 5 * $attempt))
                continue
            }
            throw
        }
    }
}

function Invoke-MgGraphBatchPages {
    param(
        [Parameter(Mandatory)]
        [array]$AllRequests,
        [int]$Size = 20,
        [int]$DelayMs = 200,
        [string]$Activity = 'Graph batch'
    )

    $allResponses = [System.Collections.Generic.List[object]]::new()
    $totalBatches = [Math]::Ceiling($AllRequests.Count / [double]$Size)

    for ($i = 0; $i -lt $AllRequests.Count; $i += $Size) {
        $batchNum = [int]($i / $Size) + 1
        $end = [Math]::Min($i + $Size - 1, $AllRequests.Count - 1)
        $chunk = @($AllRequests[$i..$end])

        Write-Progress -Activity $Activity -Status "Batch $batchNum of $totalBatches" -PercentComplete (($batchNum / $totalBatches) * 100)
        $allResponses.AddRange((Invoke-MgGraphBatch -Requests $chunk))

        if ($DelayMs -gt 0 -and ($i + $Size) -lt $AllRequests.Count) {
            Start-Sleep -Milliseconds $DelayMs
        }
    }

    Write-Progress -Activity $Activity -Completed
    return $allResponses
}

function Import-GroupCacheBatch {
    param(
        [string[]]$GroupIds,
        [hashtable]$GroupCache,
        [int]$Size,
        [int]$DelayMs
    )

    $missing = @($GroupIds | Where-Object { $_ -and -not $GroupCache.ContainsKey($_) } | Select-Object -Unique)
    if ($missing.Count -eq 0) { return }

    Write-Host "  Resolving $($missing.Count) assigning group(s)..."
    $requests = for ($i = 0; $i -lt $missing.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/groups/$($missing[$i])?`$select=id,displayName,mailNickname" }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Resolving groups'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $groupId = $missing[$index]
        if ($response.status -eq 200) {
            $GroupCache[$groupId] = [PSCustomObject]@{
                Id           = $response.body.id
                DisplayName  = $response.body.displayName
                MailNickname = $response.body.mailNickname
            }
        }
        else {
            $GroupCache[$groupId] = [PSCustomObject]@{
                Id           = $groupId
                DisplayName  = $groupId
                MailNickname = ''
            }
        }
    }
}

function Import-UserManagersBatch {
    param(
        [string[]]$UserIds,
        [hashtable]$ManagerCache,
        [int]$Size,
        [int]$DelayMs
    )

    $pending = @($UserIds | Where-Object { $_ -and -not $ManagerCache.ContainsKey($_) })
    if ($pending.Count -eq 0) { return }

    Write-Host "  Loading managers for $($pending.Count) user(s) via batch..."
    $requests = for ($i = 0; $i -lt $pending.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/users/$($pending[$i])/manager?`$select=id,displayName,userPrincipalName" }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Loading managers'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $userId = $pending[$index]
        if ($response.status -eq 200 -and $response.body.id) {
            $ManagerCache[$userId] = [PSCustomObject]@{
                Name = $response.body.displayName
                UPN  = $response.body.userPrincipalName
            }
        }
        else {
            $ManagerCache[$userId] = [PSCustomObject]@{ Name = ''; UPN = '' }
        }
    }
}

function Import-UserLicenseDetailsBatch {
    param(
        [string[]]$UserIds,
        [hashtable]$LicenseDetailCache,
        [int]$Size,
        [int]$DelayMs
    )

    $pending = @($UserIds | Where-Object { $_ -and -not $LicenseDetailCache.ContainsKey($_) })
    if ($pending.Count -eq 0) { return }

    Write-Host "  Loading license details for $($pending.Count) user(s) via batch..."
    $requests = for ($i = 0; $i -lt $pending.Count; $i++) {
        @{ id = "$i"; method = 'GET'; url = "/users/$($pending[$i])/licenseDetails" }
    }

    $responses = Invoke-MgGraphBatchPages -AllRequests $requests -Size $Size -DelayMs $DelayMs -Activity 'Loading license details'
    foreach ($response in $responses) {
        $index = [int]$response.id
        $userId = $pending[$index]
        if ($response.status -eq 200) {
            $LicenseDetailCache[$userId] = @($response.body.value)
        }
        else {
            $LicenseDetailCache[$userId] = @()
        }
    }
}

function Get-SkuInfoFromCatalog {
    param(
        [string]$SkuId,
        [hashtable]$SkuCatalog
    )

    $partNumber = $SkuCatalog[$SkuId]
    if (-not $partNumber) { $partNumber = $SkuId }
    return @{
        SkuPartNumber = $partNumber
        FriendlyName  = Get-LicenseFriendlyName -SkuPartNumber $partNumber
    }
}

function Get-LicenseAssignmentForSku {
    param(
        $LicenseAssignmentStates,
        [string]$SkuId,
        [hashtable]$GroupCache
    )

    $statesForSku = @($LicenseAssignmentStates | Where-Object { $_.SkuId.ToString() -eq $SkuId })
    $state = $statesForSku | Select-Object -First 1
    $groupNames = [System.Collections.Generic.List[string]]::new()
    $groupIds = [System.Collections.Generic.List[string]]::new()
    $hasDirect = $false

    foreach ($skuState in $statesForSku) {
        $groupInfo = Resolve-AssigningGroup -GroupId $skuState.AssignedByGroup -GroupCache $GroupCache
        if ($groupInfo.Method -eq 'Group') {
            $groupNames.Add($groupInfo.Name)
            $groupIds.Add($groupInfo.Id)
        }
        else {
            $hasDirect = $true
        }
    }

    $method = if ($groupNames.Count -gt 0 -and $hasDirect) { 'Mixed' }
    elseif ($groupNames.Count -gt 0) { 'Group' }
    elseif ($hasDirect) { 'Direct' }
    else { 'Unknown' }

    return [PSCustomObject]@{
        AssignmentMethod = $method
        GroupNames       = $groupNames
        GroupIds         = $groupIds
        AssignmentState  = if ($state) { $state.State } else { 'Unknown' }
        AssignmentError  = ($statesForSku | Where-Object { $_.Error } | ForEach-Object { $_.Error }) -join '; '
    }
}

function Get-LicenseFriendlyName {
    param([string]$SkuPartNumber)

    $names = @{
        'ENTERPRISEPACK'           = 'Microsoft 365 E3'
        'ENTERPRISEPREMIUM'        = 'Microsoft 365 E5'
        'SPE_E3'                   = 'Microsoft 365 E3'
        'SPE_E5'                   = 'Microsoft 365 E5'
        'SPE_F1'                   = 'Microsoft 365 F3'
        'M365_F1'                  = 'Microsoft 365 F1'
        'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
        'O365_BUSINESS_PREMIUM'    = 'Microsoft 365 Business Standard'
        'SPB'                      = 'Microsoft 365 Business Premium'
        'SMB_BUSINESS_PREMIUM'     = 'Microsoft 365 Business Premium'
        'EXCHANGESTANDARD'         = 'Exchange Online Plan 1'
        'EXCHANGEENTERPRISE'       = 'Exchange Online Plan 2'
        'SHAREPOINTSTANDARD'       = 'SharePoint Online Plan 1'
        'SHAREPOINTENTERPRISE'     = 'SharePoint Online Plan 2'
        'PROJECTPROFESSIONAL'      = 'Project Plan 3'
        'PROJECTPREMIUM'           = 'Project Plan 5'
        'VISIOCLIENT'              = 'Visio Plan 2'
        'POWER_BI_PRO'             = 'Power BI Pro'
        'POWER_BI_STANDARD'        = 'Power BI Free'
        'FLOW_FREE'                = 'Power Automate Free'
        'TEAMS_EXPLORATORY'        = 'Microsoft Teams Exploratory'
        'MCOPSTN2'                 = 'Microsoft Teams Phone Standard'
        'MCOEV'                    = 'Microsoft Teams Phone'
        'Microsoft_365_Copilot'    = 'Microsoft 365 Copilot'
        'WINDOWS_STORE'            = 'Windows Store for Business'
        'RIGHTSMANAGEMENT'         = 'Azure Information Protection Plan 1'
        'EMS'                      = 'Enterprise Mobility + Security E3'
        'EMSPREMIUM'               = 'Enterprise Mobility + Security E5'
        'AAD_PREMIUM'              = 'Microsoft Entra ID P1'
        'AAD_PREMIUM_P2'           = 'Microsoft Entra ID P2'
        'ATP_ENTERPRISE'           = 'Microsoft Defender for Office 365 Plan 2'
        'THREAT_INTELLIGENCE'      = 'Microsoft Defender for Office 365 Plan 2'
        'DEVELOPERPACK_E5'         = 'Microsoft 365 E5 Developer'
    }

    if ($names.ContainsKey($SkuPartNumber)) {
        return $names[$SkuPartNumber]
    }
    return $SkuPartNumber
}

function Get-InactiveDays {
    param($DateTime)

    if (-not $DateTime) { return $null }
    return [int](New-TimeSpan -Start $DateTime -End (Get-Date)).TotalDays
}

function Get-SignInActivityInfo {
    param($SignInActivity)

    $interactive = $SignInActivity.LastSignInDateTime
    $nonInteractive = $SignInActivity.LastNonInteractiveSignInDateTime

    return [PSCustomObject]@{
        LastInteractiveSignInDateTime    = $interactive
        LastNonInteractiveSignInDateTime = $nonInteractive
        InactiveDaysInteractiveSignIn    = Get-InactiveDays -DateTime $interactive
        InactiveDaysNonInteractiveSignIn = Get-InactiveDays -DateTime $nonInteractive
    }
}

function Resolve-AssigningGroup {
    param(
        $GroupId,
        [hashtable]$GroupCache
    )

    if (-not $GroupId) {
        return @{
            Method = 'Direct'
            Name   = ''
            Id     = ''
        }
    }

    $id = $GroupId.ToString()
    if (-not $GroupCache.ContainsKey($id)) {
        try {
            $group = Get-MgGroup -GroupId $id -Property Id, DisplayName, MailNickname
            $GroupCache[$id] = $group
        }
        catch {
            $GroupCache[$id] = [PSCustomObject]@{
                Id           = $id
                DisplayName  = $id
                MailNickname = ''
            }
        }
    }

    return @{
        Method = 'Group'
        Name   = $GroupCache[$id].DisplayName
        Id     = $id
    }
}

function Get-UserManagerInfo {
    param(
        [string]$UserId,
        [hashtable]$ManagerCache
    )

    if ($ManagerCache.ContainsKey($UserId)) {
        return $ManagerCache[$UserId]
    }

    return [PSCustomObject]@{ Name = ''; UPN = '' }
}

function Get-AssignmentStateSummary {
    param(
        $LicenseAssignmentStates,
        [hashtable]$SkuCatalog,
        [hashtable]$GroupCache
    )

    if (-not $LicenseAssignmentStates) {
        return '', '', 'None', @()
    }

    $licensingGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $hasDirect = $false
    $states = foreach ($state in $LicenseAssignmentStates) {
        $skuName = $SkuCatalog[$state.SkuId.ToString()]
        if (-not $skuName) { $skuName = $state.SkuId }

        $groupInfo = Resolve-AssigningGroup -GroupId $state.AssignedByGroup -GroupCache $GroupCache
        if ($groupInfo.Method -eq 'Group') {
            [void]$licensingGroups.Add($groupInfo.Name)
            $assignedBy = "Group:$($groupInfo.Name)"
        }
        else {
            $hasDirect = $true
            $assignedBy = 'Direct'
        }

        "$skuName=$($state.State)($assignedBy)"
    }

    $errors = ($LicenseAssignmentStates | Where-Object { $_.State -eq 'ActiveWithError' -and $_.Error } |
        ForEach-Object { "$($SkuCatalog[$_.SkuId.ToString()]):$($_.Error)" }) -join '; '

    $method = if ($licensingGroups.Count -gt 0 -and $hasDirect) {
        'Mixed'
    }
    elseif ($licensingGroups.Count -gt 0) {
        'Group'
    }
    elseif ($hasDirect) {
        'Direct'
    }
    else {
        'None'
    }

    return ($states -join '; '), $errors, $method, @($licensingGroups)
}

Connect-MgGraphApp

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$tenantSummaryPath = "${ExportPath}-TenantSummary-${timestamp}.csv"
$groupLicensesPath = "${ExportPath}-GroupLicenses-${timestamp}.csv"
$userSummaryPath = "${ExportPath}-Users-${timestamp}.csv"
$licenseDetailPath = "${ExportPath}-LicenseDetail-${timestamp}.csv"

$groupCache = @{}
$managerCache = @{}

Write-Host "Building tenant license catalog..."
$subscribedSkus = @(Get-MgSubscribedSku -All)
$skuCatalog = @{}
$tenantSummary = [System.Collections.Generic.List[object]]::new()

foreach ($sku in $subscribedSkus) {
    $purchased = [int]$sku.PrepaidUnits.Enabled
    $consumed = [int]$sku.ConsumedUnits
    $available = [Math]::Max(0, $purchased - $consumed)
    $friendlyName = Get-LicenseFriendlyName -SkuPartNumber $sku.SkuPartNumber

    $skuCatalog[$sku.SkuId.ToString()] = $sku.SkuPartNumber

    $tenantSummary.Add([PSCustomObject]@{
            SkuPartNumber     = $sku.SkuPartNumber
            FriendlyName      = $friendlyName
            SkuId             = $sku.SkuId
            PurchasedLicenses = $purchased
            ConsumedLicenses  = $consumed
            AvailableLicenses = $available
            CapabilityStatus  = $sku.CapabilityStatus
            AppliesTo         = $sku.AppliesTo
        })
}

Write-Host "Retrieving groups with assigned licenses..."
$licensedGroups = @(
    Get-MgGroup -All -Property Id, DisplayName, MailNickname, AssignedLicenses, LicenseProcessingState |
    Where-Object { @($_.AssignedLicenses).Count -gt 0 }
)
$groupLicenses = [System.Collections.Generic.List[object]]::new()

foreach ($group in $licensedGroups) {
    $groupCache[$group.Id] = $group

    $groupSkuPartNumbers = [System.Collections.Generic.List[string]]::new()
    $groupFriendlyNames = [System.Collections.Generic.List[string]]::new()

    foreach ($assignedLicense in $group.AssignedLicenses) {
        $partNumber = $skuCatalog[$assignedLicense.SkuId.ToString()]
        if (-not $partNumber) { $partNumber = $assignedLicense.SkuId.ToString() }
        $groupSkuPartNumbers.Add($partNumber)
        $groupFriendlyNames.Add((Get-LicenseFriendlyName -SkuPartNumber $partNumber))
    }

    $memberCount = if ($SkipGroupMemberCount) { $null } else { @(Get-MgGroupMember -GroupId $group.Id -All).Count }

    $groupLicenses.Add([PSCustomObject]@{
            GroupDisplayName       = $group.DisplayName
            GroupMailNickname      = $group.MailNickname
            GroupId                = $group.Id
            AssignedLicenses       = $groupFriendlyNames -join '; '
            AssignedSkuPartNumbers = $groupSkuPartNumbers -join '; '
            MemberCount            = $memberCount
            LicenseProcessingState = $group.LicenseProcessingState
        })
}

Write-Host "Found $($licensedGroups.Count) licensing group(s)."

$userFilter = switch ($UserType) {
    'Member' { "userType eq 'Member'" }
    'Guest' { "userType eq 'Guest'" }
    default { $null }
}

if ($DomainFilter) {
    $domainSuffix = if ($DomainFilter.StartsWith('@')) { $DomainFilter } else { "@$DomainFilter" }
    $domainClause = "endsWith(userPrincipalName, '$domainSuffix')"
    $userFilter = if ($userFilter) { "$userFilter and $domainClause" } else { $domainClause }
}

$userProperties = @(
    'Id', 'UserPrincipalName', 'DisplayName', 'Mail', 'UserType', 'AccountEnabled',
    'Department', 'JobTitle', 'CompanyName', 'CreatedDateTime',
    'AssignedLicenses', 'LicenseAssignmentStates', 'SignInActivity'
)

Write-Host "Retrieving users..."
$getUserParams = @{
    All      = $true
    Property = $userProperties
}
if ($userFilter) {
    $getUserParams['Filter'] = $userFilter
    $getUserParams['ConsistencyLevel'] = 'eventual'
    $getUserParams['CountVariable'] = 'userCount'
    Write-Host "  Filter: $userFilter"
}

$users = @(Get-MgUser @getUserParams)
Write-Host "Retrieved $($users.Count) user(s)."

# Apply switches before bulk prefetch to reduce batch volume
if ($LicensedUsersOnly) { $users = @($users | Where-Object { @($_.AssignedLicenses).Count -gt 0 }) }
if ($UnlicensedUsersOnly) { $users = @($users | Where-Object { @($_.AssignedLicenses).Count -eq 0 }) }
if ($EnabledUsersOnly) { $users = @($users | Where-Object { $_.AccountEnabled }) }
if ($DisabledUsersOnly) { $users = @($users | Where-Object { -not $_.AccountEnabled }) }
Write-Host "Users after filters: $($users.Count)"

$assigningGroupIds = @(
    $users | ForEach-Object { $_.LicenseAssignmentStates } |
    Where-Object { $_.AssignedByGroup } |
    ForEach-Object { $_.AssignedByGroup.ToString() } |
    Select-Object -Unique
)
Import-GroupCacheBatch -GroupIds $assigningGroupIds -GroupCache $groupCache -Size $BatchSize -DelayMs $BatchDelayMs

if (-not $SkipManager) {
    Import-UserManagersBatch -UserIds @($users.Id) -ManagerCache $managerCache -Size $BatchSize -DelayMs $BatchDelayMs
}

$licenseDetailCache = @{}
$licensedUserIds = @($users | Where-Object { @($_.AssignedLicenses).Count -gt 0 } | ForEach-Object { $_.Id })
if ($IncludeServicePlanDetail -and $licensedUserIds.Count -gt 0) {
    Import-UserLicenseDetailsBatch -UserIds $licensedUserIds -LicenseDetailCache $licenseDetailCache -Size $BatchSize -DelayMs $BatchDelayMs
}

Write-Host "Building report rows..."
$userSummary = [System.Collections.Generic.List[object]]::new()
$licenseDetail = [System.Collections.Generic.List[object]]::new()
$processed = 0

foreach ($user in $users) {
    $processed++
    if ($processed % 500 -eq 0 -or $processed -eq $users.Count) {
        Write-Progress -Activity 'Building report' -Status "$processed / $($users.Count)" -PercentComplete (($processed / $users.Count) * 100)
    }

    $upn = $user.UserPrincipalName

    $licenseCount = @($user.AssignedLicenses).Count
    $accountStatus = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }

    $manager = if ($SkipManager) {
        [PSCustomObject]@{ Name = ''; UPN = '' }
    }
    else {
        Get-UserManagerInfo -UserId $user.Id -ManagerCache $managerCache
    }
    $signIn = Get-SignInActivityInfo -SignInActivity $user.SignInActivity

    $assignmentSummary, $assignmentErrors, $assignmentMethod, $licensingGroups = Get-AssignmentStateSummary `
        -LicenseAssignmentStates $user.LicenseAssignmentStates `
        -SkuCatalog $skuCatalog `
        -GroupCache $groupCache

    $skuPartNumbers = [System.Collections.Generic.List[string]]::new()
    $friendlyNames = [System.Collections.Generic.List[string]]::new()
    $disabledPlansAll = [System.Collections.Generic.List[string]]::new()

    if ($licenseCount -gt 0) {
        $userLicenseDetails = if ($IncludeServicePlanDetail) { @($licenseDetailCache[$user.Id]) } else { @() }
        $detailsBySku = @{}
        foreach ($detail in $userLicenseDetails) {
            $detailsBySku[$detail.skuId.ToString()] = $detail
        }

        foreach ($assigned in $user.AssignedLicenses) {
            $skuId = $assigned.SkuId.ToString()
            $skuInfo = Get-SkuInfoFromCatalog -SkuId $skuId -SkuCatalog $skuCatalog
            $partNumber = $skuInfo.SkuPartNumber
            $friendlyName = $skuInfo.FriendlyName
            $skuPartNumbers.Add($partNumber)
            $friendlyNames.Add($friendlyName)

            $disabledPlanText = ''
            if ($IncludeServicePlanDetail -and $detailsBySku.ContainsKey($skuId)) {
                $disabledPlans = @(
                    $detailsBySku[$skuId].servicePlans |
                    Where-Object { $_.provisioningStatus -eq 'Disabled' } |
                    ForEach-Object { $_.servicePlanName }
                )
                $disabledPlanText = $disabledPlans -join '; '
                if ($disabledPlans.Count -gt 0) {
                    $disabledPlansAll.Add("$partNumber=[$disabledPlanText]")
                }
            }

            $assignment = Get-LicenseAssignmentForSku `
                -LicenseAssignmentStates $user.LicenseAssignmentStates `
                -SkuId $skuId `
                -GroupCache $groupCache

            $licenseDetail.Add([PSCustomObject]@{
                    UserPrincipalName                = $upn
                    DisplayName                      = $user.DisplayName
                    Manager                          = $manager.Name
                    ManagerUPN                       = $manager.UPN
                    LastInteractiveSignInDateTime    = $signIn.LastInteractiveSignInDateTime
                    LastNonInteractiveSignInDateTime = $signIn.LastNonInteractiveSignInDateTime
                    InactiveDaysInteractiveSignIn    = $signIn.InactiveDaysInteractiveSignIn
                    InactiveDaysNonInteractiveSignIn = $signIn.InactiveDaysNonInteractiveSignIn
                    SkuPartNumber                    = $partNumber
                    FriendlyName                     = $friendlyName
                    SkuId                            = $assigned.SkuId
                    AssignmentMethod                 = $assignment.AssignmentMethod
                    AssigningGroupName               = ($assignment.GroupNames -join '; ')
                    AssigningGroupId                 = ($assignment.GroupIds -join '; ')
                    AssignmentState                  = $assignment.AssignmentState
                    AssignmentError                  = $assignment.AssignmentError
                    DisabledServicePlans             = $disabledPlanText
                })
        }
    }

    $userSummary.Add([PSCustomObject]@{
            UserPrincipalName                = $upn
            DisplayName                      = $user.DisplayName
            Mail                             = $user.Mail
            Manager                          = $manager.Name
            ManagerUPN                       = $manager.UPN
            UserType                         = $user.UserType
            AccountStatus                    = $accountStatus
            Department                       = $user.Department
            JobTitle                         = $user.JobTitle
            CompanyName                      = $user.CompanyName
            CreatedDateTime                  = $user.CreatedDateTime
            LicenseCount                     = $licenseCount
            AssignedLicenses                 = if ($friendlyNames.Count -gt 0) { $friendlyNames -join '; ' } else { 'No License Assigned' }
            AssignedSkuPartNumbers           = if ($skuPartNumbers.Count -gt 0) { $skuPartNumbers -join '; ' } else { '' }
            DisabledServicePlans             = ($disabledPlansAll -join ' | ')
            LicenseAssignmentMethod          = $assignmentMethod
            LicensingGroups                  = ($licensingGroups -join '; ')
            LicenseAssignmentStates          = $assignmentSummary
            LicenseAssignmentErrors          = $assignmentErrors
            LastInteractiveSignInDateTime    = $signIn.LastInteractiveSignInDateTime
            LastNonInteractiveSignInDateTime = $signIn.LastNonInteractiveSignInDateTime
            InactiveDaysInteractiveSignIn    = $signIn.InactiveDaysInteractiveSignIn
            InactiveDaysNonInteractiveSignIn = $signIn.InactiveDaysNonInteractiveSignIn
            ObjectId                         = $user.Id
        })
}

Write-Progress -Activity 'Building report' -Completed

$tenantSummary | Export-Csv -LiteralPath $tenantSummaryPath -NoTypeInformation -Encoding UTF8
$groupLicenses | Export-Csv -LiteralPath $groupLicensesPath -NoTypeInformation -Encoding UTF8
$userSummary | Export-Csv -LiteralPath $userSummaryPath -NoTypeInformation -Encoding UTF8
$licenseDetail | Export-Csv -LiteralPath $licenseDetailPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Report complete:"
Write-Host "  Tenant SKUs    : $($tenantSummary.Count) rows -> $tenantSummaryPath"
Write-Host "  Group licenses : $($groupLicenses.Count) rows -> $groupLicensesPath"
Write-Host "  Users          : $($userSummary.Count) rows -> $userSummaryPath"
Write-Host "  License detail : $($licenseDetail.Count) rows -> $licenseDetailPath"
