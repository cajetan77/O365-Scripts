param(
    [Parameter(Mandatory = $false)]
    [string]$EnvironmentName = "cajesharepoint (default)",

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = ".\UserSecurityRoles.csv",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemApplicationUsers
)

Import-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -ErrorAction Stop
Import-Module -Name Microsoft.Xrm.Data.PowerShell -ErrorAction Stop

function Expand-CrmOrganizations {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RawOrganizations
    )

    if ($null -eq $RawOrganizations) {
        return @()
    }

    if ($null -ne $RawOrganizations.Organizations) {
        return @($RawOrganizations.Organizations)
    }

    $asArray = @($RawOrganizations)
    if ($asArray.Count -gt 1) {
        return $asArray
    }

    $single = $asArray[0]
    $names = @($single.FriendlyName)
    if ($names.Count -le 1) {
        return @($single)
    }

    $urlProps = @("WebApplicationUrl", "Url", "OrganizationUrl", "OrganizationServiceUrl")
    $urls = @()
    foreach ($prop in $urlProps) {
        $value = $single.$prop
        if ($null -eq $value) { continue }
        $candidateUrls = @($value)
        if ($candidateUrls.Count -ge $names.Count) {
            $urls = $candidateUrls
            break
        }
    }

    $uniqueNames = @($single.UniqueName)
    if ($uniqueNames.Count -eq 1 -and $names.Count -gt 1) {
        $uniqueNames = @()
    }

    $expanded = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $names.Count; $i++) {
        $url = if ($i -lt $urls.Count) { $urls[$i] } else { $null }
        $uniqueName = if ($i -lt $uniqueNames.Count) { $uniqueNames[$i] } else { $null }

        $expanded.Add([PSCustomObject]@{
                FriendlyName           = [string]$names[$i]
                UniqueName             = $uniqueName
                WebApplicationUrl      = $url
                Url                    = $url
                OrganizationUrl        = $url
                OrganizationServiceUrl = $url
            })
    }

    return $expanded.ToArray()
}

function Resolve-OrganizationUrl {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Organization
    )

    $urlCandidates = @(
        $Organization.WebApplicationUrl,
        $Organization.Url,
        $Organization.OrganizationUrl,
        $Organization.OrganizationServiceUrl
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($urlCandidates.Count -gt 0) {
        return $urlCandidates[0]
    }

    return $null
}

Write-Host "Enter credentials for Power Platform / Dataverse..."
$credential = Get-Credential

$orgsRaw = Get-CrmOrganizations -Credential $credential -OnlineType Office365
$orgs = @(Expand-CrmOrganizations -RawOrganizations $orgsRaw)
if ($orgs.Count -eq 0) {
    throw "No Dataverse environments were returned for this account."
}

$targetOrg = $orgs |
Where-Object { $_.FriendlyName -eq $EnvironmentName -or $_.UniqueName -eq $EnvironmentName } |
Select-Object -First 1

if (-not $targetOrg) {
    $available = ($orgs | ForEach-Object { $_.FriendlyName } | Sort-Object -Unique) -join ", "
    throw "Environment '$EnvironmentName' was not found. Available environments: $available"
}

$serverUrl = Resolve-OrganizationUrl -Organization $targetOrg
if ([string]::IsNullOrWhiteSpace($serverUrl)) {
    throw "Could not resolve URL for environment '$EnvironmentName'."
}

Write-Host "Connecting to environment: $($targetOrg.FriendlyName)"
$connection = Connect-CrmOnline -Credential $credential -ServerUrl $serverUrl
if (-not $connection -or -not $connection.IsReady) {
    throw "Connection failed for environment '$EnvironmentName'."
}

$fetchXml = @"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='systemuser'>
    <attribute name='systemuserid' />
    <attribute name='fullname' />
    <attribute name='internalemailaddress' />
    <attribute name='domainname' />
    <attribute name='isdisabled' />
    <attribute name='accessmode' />
    <attribute name='applicationid' />
    <attribute name='businessunitid' />
    <order attribute='fullname' descending='false' />
    <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' intersect='true'>
      <link-entity name='role' from='roleid' to='roleid' alias='r'>
        <attribute name='name' />
        <attribute name='roleid' />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

Write-Host "Fetching user security role assignments..."
$results = Get-CrmRecordsByFetch -Conn $connection -Fetch $fetchXml -AllRows
$records = @($results.CrmRecords)

if (-not $IncludeSystemApplicationUsers) {
    $records = @(
        $records | Where-Object {
            $isSystemName = $_.fullname -eq "SYSTEM"
            $isApplicationUser = -not [string]::IsNullOrWhiteSpace([string]$_.applicationid)
            $accessModeText = [string]$_.accessmode
            $accessModeValue = 0
            $hasNumericAccessMode = [int]::TryParse($accessModeText, [ref]$accessModeValue)
            $isNonInteractive = $false

            if ($hasNumericAccessMode) {
                $isNonInteractive = $accessModeValue -eq 4
            }
            elseif (-not [string]::IsNullOrWhiteSpace($accessModeText)) {
                # Some tenants return friendly labels (for example "Read-Write")
                # instead of integer enum values.
                $isNonInteractive = $accessModeText -match "non[-\s]?interactive"
            }

            -not ($isSystemName -or $isApplicationUser -or $isNonInteractive)
        }
    )
}

$rows = foreach ($record in $records) {
    [PSCustomObject]@{
        EnvironmentName = $targetOrg.FriendlyName
        EnvironmentUrl  = $serverUrl
        UserId          = $record.systemuserid
        FullName        = $record.fullname
        EmailAddress    = $record.internalemailaddress
        DomainName      = $record.domainname
        IsDisabled      = $record.isdisabled
        AccessMode      = $record.accessmode
        ApplicationId   = $record.applicationid
        BusinessUnit    = $record.businessunitid
        RoleId          = $record.'r.roleid'
        RoleName        = $record.'r.name'
    }
}

$rows |
Sort-Object FullName, RoleName |
Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "Export completed: $OutputFile"
Write-Host "Total user-role records: $($rows.Count)"
