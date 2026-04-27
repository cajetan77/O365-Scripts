Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell
Install-Module -Name Microsoft.Xrm.Data.PowerShell -AllowClobber

$cred = Get-Credential

<#
  Get-CrmOrganizations sometimes returns ONE object whose properties (FriendlyName, Url, …)
  are parallel arrays — one entry per Dataverse org. Wrapping that in @() yields a single
  element, so foreach would run once and "$($org.FriendlyName)" becomes all names joined.
  This helper expands either pattern into one PSCustomObject per environment.
#>
function Expand-CrmOrganizations {
  param([object] $Raw)

  if ($null -eq $Raw) { return @() }

  if ($null -ne $Raw.Organizations) {
    return @($Raw.Organizations)
  }

  $asArray = @($Raw)
  if ($asArray.Count -gt 1) { return $asArray }

  $single = $asArray[0]
  $names = @($single.FriendlyName)
  if ($names.Count -le 1) { return @($single) }

  $urlProps = @(
    'WebApplicationUrl',
    'Url',
    'OrganizationUrl',
    'OrganizationServiceUrl'
  )
  $urls = @()
  foreach ($p in $urlProps) {
    $v = $single.$p
    if ($null -eq $v) { continue }
    $a = @($v)
    if ($a.Count -ge $names.Count) {
      $urls = $a
      break
    }
  }

  $uniqueNames = @($single.UniqueName)
  if ($uniqueNames.Count -eq 1 -and $names.Count -gt 1) { $uniqueNames = @() }

  $out = [System.Collections.Generic.List[object]]::new()
  for ($i = 0; $i -lt $names.Count; $i++) {
    $u = if ($i -lt $urls.Count) { $urls[$i] } else { $null }
    $un = if ($i -lt $uniqueNames.Count) { $uniqueNames[$i] } else { $null }
    $null = $out.Add([PSCustomObject]@{
        FriendlyName        = [string]$names[$i]
        UniqueName          = $un
        WebApplicationUrl   = $u
        Url                 = $u
        OrganizationUrl     = $u
      })
  }
  return $out.ToArray()
}

$crmOrgsRaw = Get-CrmOrganizations -Credential $cred -OnlineType Office365
$crmOrgs = @(Expand-CrmOrganizations -Raw $crmOrgsRaw)
$allComponents = [System.Collections.Generic.List[object]]::new()
$outputFile = ".\All-Solution-Components.csv"
$totalOrgs = $crmOrgs.Count
$orgIndex = 0

# Include solution metadata so each component can be traced to its parent solution.
$fetchXml = @"
<fetch>
  <entity name="solutioncomponent">
    <attribute name="solutioncomponentid" />
    <attribute name="componenttype" />
    <attribute name="objectid" />
    <attribute name="solutionid" />
    <attribute name="createdon" />
    <attribute name="modifiedon" />
    <attribute name="ismetadata" />
    <attribute name="rootsolutioncomponentid" />
    <link-entity name="solution" from="solutionid" to="solutionid" alias="sol">
      <attribute name="friendlyname" />
      <attribute name="uniquename" />
      <attribute name="version" />
      <link-entity name="publisher" from="publisherid" to="publisherid" alias="pub">
        <attribute name="publisherid" />
        <attribute name="friendlyname" />
        <attribute name="uniquename" />
        <attribute name="customizationprefix" />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

function Get-OverallPercentComplete {
  param(
    [int] $OrgIndex,
    [int] $TotalOrgs,
    [int] $RowIndex,
    [int] $RowTotal
  )
  if ($TotalOrgs -lt 1) { return 100 }
  $orgFraction = if ($RowTotal -gt 0) { [double]$RowIndex / $RowTotal } else { 1.0 }
  $value = (($OrgIndex - 1) + $orgFraction) / $TotalOrgs * 100.0
  return [int][math]::Min(100, [math]::Max(0, [math]::Round($value)))
}

foreach ($org in $crmOrgs) {
  $orgIndex++

  $serverUrl = $org.WebApplicationUrl
  if ([string]::IsNullOrWhiteSpace($serverUrl)) { $serverUrl = $org.Url }
  if ([string]::IsNullOrWhiteSpace($serverUrl)) {
    $serverUrl = $org.OrganizationUrl
  }
  if ([string]::IsNullOrWhiteSpace($serverUrl)) {
    Write-Warning "Skipping org because URL was not found: $($org | Out-String)"
    Write-Progress -Id 1 -Activity "Exporting solution components" `
      -Status "Environment $orgIndex of $totalOrgs (skipped)" `
      -CurrentOperation $org.FriendlyName `
      -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex 1 -RowTotal 1)
    continue
  }

  Write-Host "Processing environment: $($org.FriendlyName) [$serverUrl]"

  Write-Progress -Id 1 -Activity "Exporting solution components" `
    -Status "Environment $orgIndex of $totalOrgs" `
    -CurrentOperation "Connecting: $($org.FriendlyName)" `
    -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex 0 -RowTotal 1)

  $connection = Connect-CrmOnline -Credential $cred -ServerUrl $serverUrl
  if (-not $connection -or -not $connection.IsReady) {
    Write-Warning "Could not connect to environment: $($org.FriendlyName)"
    Write-Progress -Id 1 -Activity "Exporting solution components" `
      -Status "Environment $orgIndex of $totalOrgs" `
      -CurrentOperation "Failed: $($org.FriendlyName)" `
      -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex 1 -RowTotal 1)
    continue
  }

  Write-Progress -Id 1 -Activity "Exporting solution components" `
    -Status "Environment $orgIndex of $totalOrgs" `
    -CurrentOperation "Fetching components: $($org.FriendlyName)" `
    -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex 0 -RowTotal 1)

  $results = Get-CrmRecordsByFetch -Conn $connection -Fetch $fetchXml -AllRows

  $records = @($results.CrmRecords)
  $recTotal = $records.Count
  $updateEvery = [math]::Max(1, [int][math]::Ceiling($recTotal / 150.0))
  $recNum = 0

  foreach ($item in $records) {
    $recNum++
    if (($recNum % $updateEvery -eq 0) -or ($recNum -eq $recTotal)) {
      Write-Progress -Id 1 -Activity "Exporting solution components" `
        -Status "Environment $orgIndex of $totalOrgs" `
        -CurrentOperation "Building rows: $recNum / $recTotal ($($org.FriendlyName))" `
        -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex $recNum -RowTotal $recTotal)
    }

    $allComponents.Add([PSCustomObject]@{
      EnvironmentName       = $org.FriendlyName
      EnvironmentUrl        = $serverUrl
      SolutionId            = $item.solutionid
      SolutionUniqueName    = $item.'sol.uniquename'
      SolutionFriendlyName  = $item.'sol.friendlyname'
      SolutionVersion       = $item.'sol.version'
      PublisherId           = $item.'pub.publisherid'
      PublisherFriendlyName   = $item.'pub.friendlyname'
      PublisherUniqueName     = $item.'pub.uniquename'
      PublisherPrefix         = $item.'pub.customizationprefix'
      SolutionComponentId   = $item.solutioncomponentid
      RootSolutionComponent = $item.rootsolutioncomponentid
      ComponentType         = $item.componenttype
      ObjectId              = $item.objectid
      IsMetadata            = $item.ismetadata
      CreatedOn             = $item.createdon
      ModifiedOn            = $item.modifiedon
    })
  }

  if ($recTotal -eq 0) {
    Write-Progress -Id 1 -Activity "Exporting solution components" `
      -Status "Environment $orgIndex of $totalOrgs" `
      -CurrentOperation "No components: $($org.FriendlyName)" `
      -PercentComplete (Get-OverallPercentComplete -OrgIndex $orgIndex -TotalOrgs $totalOrgs -RowIndex 1 -RowTotal 1)
  }
}

Write-Progress -Id 1 -Activity "Exporting solution components" `
  -Status "Writing CSV" `
  -CurrentOperation $outputFile `
  -PercentComplete 100

$allComponents |
  Sort-Object EnvironmentName, SolutionUniqueName, ComponentType, ObjectId |
  Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Progress -Id 1 -Activity "Exporting solution components" -Completed

Write-Host "Export complete: $outputFile"
Write-Host "Total rows exported: $($allComponents.Count)"

