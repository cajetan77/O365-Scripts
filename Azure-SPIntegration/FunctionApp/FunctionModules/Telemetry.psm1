$script:ProvisionTraceContext = @{}
$script:AppInsightsInitialized = $false

function Initialize-ProvisionTraceContext {
    param(
        [hashtable]$Dimensions
    )

    $script:ProvisionTraceContext = @{}
    foreach ($key in $Dimensions.Keys) {
        if ($null -ne $Dimensions[$key] -and -not [string]::IsNullOrWhiteSpace([string]$Dimensions[$key])) {
            $script:ProvisionTraceContext[$key] = [string]$Dimensions[$key]
        }
    }
}

function Initialize-ApplicationInsights {
    if ($script:AppInsightsInitialized) {
        return
    }

    $moduleRoot = $PSScriptRoot
    if ([string]::IsNullOrEmpty($moduleRoot)) {
        $moduleRoot = Join-Path $PWD.Path 'FunctionModules'
    }

    $dllCandidates = @(
        (Join-Path $moduleRoot '../ExternalModules/PnP.PowerShell/3.2.0/Common/Microsoft.ApplicationInsights.dll')
        (Join-Path $moduleRoot 'ExternalModules/PnP.PowerShell/3.2.0/Common/Microsoft.ApplicationInsights.dll')
    )

    foreach ($dllPath in $dllCandidates) {
        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath -ErrorAction Stop
            break
        }
    }

    $script:AppInsightsInitialized = $true
}

function Write-ProvisionTrace {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info',

        [hashtable]$Properties = @{}
    )

    $dimensions = @{}
    foreach ($key in $script:ProvisionTraceContext.Keys) {
        $dimensions[$key] = $script:ProvisionTraceContext[$key]
    }
    foreach ($key in $Properties.Keys) {
        if ($null -ne $Properties[$key]) {
            $dimensions[$key] = [string]$Properties[$key]
        }
    }

    switch ($Level) {
        'Error' { Write-Host "ERROR: $Message" }
        'Warning' { Write-Warning $Message }
        default { Write-Host $Message }
    }

    try {
        Initialize-ApplicationInsights

        $configType = [Microsoft.ApplicationInsights.Extensibility.TelemetryConfiguration]
        $clientType = [Microsoft.ApplicationInsights.TelemetryClient]
        $severityType = [Microsoft.ApplicationInsights.DataContracts.SeverityLevel]

        if (-not $configType -or -not $clientType) {
            return
        }

        $config = $configType::Active
        if (-not $config -or (
                [string]::IsNullOrWhiteSpace($config.ConnectionString) -and
                [string]::IsNullOrWhiteSpace($config.InstrumentationKey))) {
            return
        }

        $client = [Microsoft.ApplicationInsights.TelemetryClient]::new($config)
        $severity = switch ($Level) {
            'Error' { [Microsoft.ApplicationInsights.DataContracts.SeverityLevel]::Error }
            'Warning' { [Microsoft.ApplicationInsights.DataContracts.SeverityLevel]::Warning }
            default { [Microsoft.ApplicationInsights.DataContracts.SeverityLevel]::Information }
        }

        $propertyBag = [System.Collections.Generic.Dictionary[string, string]]::new()
        foreach ($key in $dimensions.Keys) {
            [void]$propertyBag.Add($key, $dimensions[$key])
        }

        $client.TrackTrace($Message, $severity, $propertyBag)
        $client.Flush()
    }
    catch {
        # Host stdout logging still captured when App Insights SDK is unavailable locally.
    }
}

Export-ModuleMember -Function Initialize-ProvisionTraceContext, Write-ProvisionTrace
