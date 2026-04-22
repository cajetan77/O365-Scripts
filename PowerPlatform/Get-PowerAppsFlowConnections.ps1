Import-Module -Name Microsoft.PowerApps.Administration.PowerShell

$outputFile = ".\powerplatinventory.csv"

$environments = Get-AdminPowerAppEnvironment | Where-Object { $_.DisplayName -eq "cajesharepoint (default)" } # Include default environment
$powerPlatObjects = @()
foreach ($e in $environments) {
    write-host "Environment: " $e.displayname
    $powerapps = Get-AdminPowerApp -EnvironmentName $e.EnvironmentName
    foreach ($pa in $powerapps) {
        write-host "  App Name: " $pa.DisplayName " - " $pa.AppName
        foreach ($conRef in $pa.Internal.properties.connectionReferences) {
            foreach ($con in $conRef) {
                foreach ($conId in ($con | Get-Member -MemberType NoteProperty).Name) {
                    $conDetails = $($con.$conId)
                    $apiTier = $conDetails.apiTier
                    if ($conDetails.isCustomApiConnection) { $apiTier = "Premium (CustomAPI)" }
                    if ($conDetails.isOnPremiseConnection ) { $apiTier = "Premium (OnPrem)" }
                    Write-Host "    " $conDetails.displayName " (" $apiTier ")"
                    $paObj = @{
                        type           = "Power App"
                        ConnectionName = $conDetails.displayName
                        ConnectionId   = $conDetails.id
                        Tier           = $apiTier
                        Environment    = $e.displayname
                        AppFlowName    = $pa.DisplayName
                        createdDate    = $pa.CreatedTime
                        createdBy      = $pa.Owner
                    }
                    $powerPlatObjects += $(new-object psobject -Property $paObj)
                } #foreach $conId
            } #foreach $con
        } #foreach $conRef
    } #foreach power app
    $flows = Get-AdminFlow -EnvironmentName $e.EnvironmentName
    foreach ($f in $flows) {
        Write-Host "  Flow Name: " $f.DisplayName " - " $f.FlowName
        $fl = get-adminflow -FlowName $f.FlowName -EnvironmentName $e.EnvironmentName
        foreach ($conRef in $fl.Internal.properties.connectionReferences) {
            foreach ($con in $conRef) {
                foreach ($conId in ($con | Get-Member -MemberType NoteProperty).Name) {
                    $conDetails = $($con.$conId)
                    $apiTier = $conDetails.apiDefinition.properties.tier
                    if ($conDetails.apiDefinition.properties.isCustomApi) { $apiTier = "Premium (CustomAPI)" }
                    Write-Host "    " $conDetails.displayName " (" $apiTier ")"
                    $paObj = @{
                        type           = "Power Automate"
                        ConnectionName = $conDetails.displayName
                        ConnectionId   = $conDetails.id
                        Tier           = $apiTier
                        Environment    = $e.DisplayName
                        AppFlowName    = $f.DisplayName
                        createdDate    = $f.CreatedTime
                        createdBy      = $f.CreatedBy
                    }
                    $powerPlatObjects += $(new-object psobject -Property $paObj)
                } #foreach $conId
            } #foreach $con
        } #foreach $conRef
    } #foreach flow
} #foreach environment
$powerPlatObjects | Export-Csv $outputFile -NoTypeInformation