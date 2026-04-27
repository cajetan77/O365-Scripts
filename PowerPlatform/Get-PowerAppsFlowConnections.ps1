Import-Module -Name Microsoft.PowerApps.Administration.PowerShell

$outputFile = ".\powerapps.csv"
$outputflowFile = ".\powerplatflows.csv"

function Get-PowerAppCategory {
    param (
        [Parameter(Mandatory = $true)]
        $PowerApp
    )

    $candidateTypes = @(
        $PowerApp.AppType,
        $PowerApp.Internal.properties.appType,
        $PowerApp.Internal.properties.powerAppType
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidate in $candidateTypes) {
        if ($candidate.ToString() -match "model") {
            return "Power App (Model-Driven)"
        }
    }

    return "Power App (Canvas)"
}

function Get-SolutionInfo {
    param (
        [Parameter(Mandatory = $true)]
        $Item
    )

    $solutionIdCandidates = @(
        $Item.Internal.properties.solutionid,
        $Item.Internal.properties.solutionId,
        $Item.Properties.solutionid,
        $Item.Properties.solutionId
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $solutionId = $null
    if ($solutionIdCandidates.Count -gt 0) {
        $solutionId = $solutionIdCandidates[0].ToString()
    }

    return @{
        IsInSolution = -not [string]::IsNullOrWhiteSpace($solutionId)
        SolutionId   = $solutionId
    }
}

Import-Module Microsoft.Xrm.Data.Powershell
$CRMOrgs = Get-CrmOrganizations -Credential $Cred -DeploymentRegion NorthAmerica –OnlineType Office365


[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Get-CrmConnection  

$environments = Get-AdminPowerAppEnvironment | Where-Object { $_.DisplayName -eq "Caje Dev" } # Include default environment
$powerAppObjects = @()
$powerPlatObjects = @()
foreach ($e in $environments) {
    write-host "Environment: " $e.displayname
    $powerapps = Get-AdminPowerApp -EnvironmentName $e.EnvironmentName  # Exclude Portal apps
    foreach ($pa in $powerapps) {
        write-host "  App Name: " $pa.DisplayName " - " $pa.AppName
        foreach ($conRef in $pa.Internal.properties.connectionReferences) {
            foreach ($con in $conRef) {
                foreach ($conId in ($con | Get-Member -MemberType NoteProperty).Name) {
                    $conDetails = $($con.$conId)
                    $apiTier = $conDetails.apiTier
                    if ($conDetails.isCustomApiConnection) { $apiTier = "Premium (CustomAPI)" }
                    if ($conDetails.isOnPremiseConnection ) { $apiTier = "Premium (OnPrem)" }
                    $appCategory = Get-PowerAppCategory -PowerApp $pa
                    $appSolutionInfo = Get-SolutionInfo -Item $pa
                    Write-Host "    " $conDetails.displayName " (" $apiTier ")"
                    $paObj = @{
                        type           = $appCategory
                        ConnectionName = $conDetails.displayName
                        ConnectionId   = $conDetails.id
                        Tier           = $apiTier
                        Environment    = $e.displayname
                        AppFlowName    = $pa.DisplayName
                        createdDate    = $pa.CreatedTime
                        createdBy      = $pa.Owner
                        IsInSolution   = $appSolutionInfo.IsInSolution
                        SolutionId     = $appSolutionInfo.SolutionId
                    }
                    $powerAppObjects += $(new-object psobject -Property $paObj)
                } #foreach $conId
            } #foreach $con
        } #foreach $conRef
    } #foreach power app
    $powerAppObjects | Export-Csv $outputFile  -NoTypeInformation
    $flows = Get-AdminFlow -EnvironmentName $e.EnvironmentName
    foreach ($f in $flows) {
        Write-Host "  Flow Name: " $f.DisplayName " - " $f.FlowName
        $fl = get-adminflow -FlowName $f.FlowName -EnvironmentName $e.EnvironmentName
        $flowSolutionInfo = Get-SolutionInfo -Item $fl
        foreach ($conRef in $fl.Internal.properties.connectionReferences) {
            foreach ($con in $conRef) {
                foreach ($conId in ($con | Get-Member -MemberType NoteProperty).Name) {
                    $conDetails = $($con.$conId)
                    $apiTier = $conDetails.apiDefinition.properties.tier
                    if ($conDetails.apiDefinition.properties.isCustomApi) { $apiTier = "Premium (CustomAPI)" }
                    Write-Host "    " $conDetails.displayName " (" $apiTier ")"
                    $flowObj = @{
                        type                 = "Power Automate"
                        ConnectionName       = $conDetails.displayName
                        ConnectionId         = $conDetails.id
                        Tier                 = $apiTier
                        Environment          = $e.DisplayName
                        AppFlowName          = $f.DisplayName
                        createdDate          = $f.CreatedTime
                        createdBy            = $f.CreatedBy
                        Enabled              = $f.Enabled
                        ModifiedDate         = $fl.Internal.properties.lastModifiedTime
                        FlowSuspensionReason = $fl.Internal.properties.flowSuspensionReason
                        IsInSolution         = $flowSolutionInfo.IsInSolution
                        SolutionId           = $flowSolutionInfo.SolutionId
                    }
                    $powerPlatObjects += $(new-object psobject -Property $flowObj)
                } #foreach $conId
            } #foreach $con
        } #foreach $conRef
    } #foreach flow
} #foreach environment
$powerPlatObjects | Export-Csv $outputflowFile  -NoTypeInformation