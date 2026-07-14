#Requires -Version 7.4
using namespace System.Net

param($Request, $TriggerMetadata)

Import-Module "$PSScriptRoot\..\FunctionModules\Telemetry.psm1" -Force
Import-Module "$PSScriptRoot\..\FunctionModules\Provisioning.psm1" -Force

function Test-SecureHeaderMatch {
    param(
        [string]$Provided,
        [string]$Expected
    )

    if ([string]::IsNullOrEmpty($Expected)) {
        return $false
    }

    $providedBytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Provided)
    $expectedBytes = [System.Text.Encoding]::UTF8.GetBytes($Expected)

    if ($providedBytes.Length -ne $expectedBytes.Length) {
        return $false
    }

    $difference = 0
    for ($index = 0; $index -lt $providedBytes.Length; $index++) {
        $difference = $difference -bor ($providedBytes[$index] -bxor $expectedBytes[$index])
    }

    return $difference -eq 0
}

try {
    $expectedInternalKey = $env:FUNCTION_HEADER_VALUE
    if ([string]::IsNullOrWhiteSpace($expectedInternalKey)) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::InternalServerError
                Body       = (@{
                        status  = 'Error'
                        message = 'FUNCTION_HEADER_VALUE is not configured.'
                    } | ConvertTo-Json)
                Headers    = @{
                    'Content-Type' = 'application/json'
                }
            })
        return
    }

    $incomingInternalKey = [string]$Request.Headers['X-INTERNAL-KEY']
    if (-not (Test-SecureHeaderMatch -Provided $incomingInternalKey -Expected $expectedInternalKey)) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::Unauthorized
                Body       = (@{
                        status  = 'Unauthorized'
                        message = 'Invalid internal key.'
                    } | ConvertTo-Json)
                Headers    = @{
                    'Content-Type' = 'application/json'
                }
            })
        return
    }
    Write-Host ($Request.Body | ConvertTo-Json)
    $siteUrl = [string]$Request.Body.ObjectUrl
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        $siteUrl = [string]$Request.Body.objectUrl
    }

    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body       = (@{
                        status  = 'BadRequest'
                        message = 'objectUrl is required.'
                    } | ConvertTo-Json)
                Headers    = @{
                    'Content-Type' = 'application/json'
                }
            })
        return
    }

    if ($siteUrl -notmatch '^https://[\w\-.]+\.sharepoint\.com/sites/[\w\-.]+(?:/.*)?$') {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body       = (@{
                        status  = 'BadRequest'
                        message = 'objectUrl must be a valid SharePoint site URL.'
                    } | ConvertTo-Json)
                Headers    = @{
                    'Content-Type' = 'application/json'
                }
            })
        return
    }

    $siteName = $siteUrl.Split('/')[-1]
    Initialize-ProvisionTraceContext -Dimensions @{
        SiteUrl      = $siteUrl
        SiteName     = $siteName
        InvocationId = $TriggerMetadata.InvocationId
        FunctionName = 'ProvisionSite'
    }

    Write-ProvisionTrace -Message 'Starting SharePoint provisioning' -Properties @{
        Step = 'Start'
    }

    $result = Invoke-SharePointProvisioning -SiteUrl $siteUrl

    Write-ProvisionTrace -Message 'SharePoint provisioning completed' -Properties @{
        Step   = 'Complete'
        Status = $result.status
    }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = ($result | ConvertTo-Json -Depth 10)
            Headers    = @{
                'Content-Type' = 'application/json'
            }
        })
}
catch {
    $failedStep = 'Unknown'
    $errorMessage = $_.Exception.Message
    $currentException = $_.Exception

    while ($null -ne $currentException) {
        if ($currentException.Message -match '\[([^\]]+)\]') {
            $failedStep = $matches[1]
            $errorMessage = $currentException.Message
            break
        }

        $currentException = $currentException.InnerException
    }

    if ($failedStep -eq 'Unknown' -and -not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
        $errorMessage = $_.ErrorDetails.Message
        if ($errorMessage -match '\[([^\]]+)\]') {
            $failedStep = $matches[1]
        }
    }

    if ($failedStep -eq 'Unknown' -and (Get-Command Get-CurrentProvisionStep -ErrorAction SilentlyContinue)) {
        $currentStep = Get-CurrentProvisionStep
        if (-not [string]::IsNullOrWhiteSpace($currentStep) -and $currentStep -ne 'Start') {
            $failedStep = $currentStep
        }
    }

    Write-ProvisionTrace -Message $errorMessage -Level Error -Properties @{
        Step       = $failedStep
        ScriptLine = $_.InvocationInfo.ScriptLineNumber
        ScriptName = $_.InvocationInfo.ScriptName
    }

    $errorBody = @{
        status  = 'Error'
        step    = $failedStep
        message = $errorMessage
    } | ConvertTo-Json

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::InternalServerError
            Body       = $errorBody
            Headers    = @{
                'Content-Type' = 'application/json'
            }
        })
}
