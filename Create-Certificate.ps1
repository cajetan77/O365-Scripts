Import-Module PSPKI -ErrorAction Stop

$subject = "SiteProvisioningCert"  # Common Name (CN) for the certificate
$yearsGood = 5  # How many years the cert will be good for.
$friendlyName = "Site Provisioning Certificate"  # Friendly name for the certificate in the store
#$pfxPass = (Get-Credential).Password
$exportTo = ".\$subject.pfx"
$cerPath = ".\$subject.cer"

Write-Host "Creating self-signed certificate..." -ForegroundColor Cyan

# Create the certificate and export to PFX
if (Test-Path $exportTo) {
    Write-Host "Certificate already exists. Skipping creation." -ForegroundColor Yellow
    
}
else {
    $cert = New-SelfSignedCertificateEx -Subject "CN=$subject" -NotAfter $((Get-Date).AddYears($yearsGood)) -AlgorithmName RSA -KeyLength 2048 -KeySpec Exchange -KeyUsage "DigitalSignature", "KeyEncipherment" -EnhancedKeyUsage "Server Authentication", "Client Authentication" -SignatureAlgorithm SHA256 -FriendlyName $friendlyName -Exportable:$true -Path $exportTo
    <# Action when all if and elseif conditions are false #>
}
#-Password $pfxPass

if (-not $cert) {
    Write-Error "Failed to create certificate"
    exit 1
}

Write-Host "Certificate created successfully. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
Write-Host "Waiting for certificate to be available in store..." -ForegroundColor Yellow

# Wait a moment for the certificate to be available in the store
Start-Sleep -Seconds 2

# Get certificate from store using thumbprint
$thumbprint = $cert.Thumbprint
$storeCert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $thumbprint }

if ($storeCert) {
    try {
        Write-Host "Exporting public certificate to .cer file..." -ForegroundColor Cyan
        
        # Export the public certificate (.cer file) - contains only the public key
        Export-Certificate -Cert $storeCert -FilePath $cerPath -Type CERT -ErrorAction Stop
        
        Write-Host ""
        Write-Host "Certificate files created successfully :" -ForegroundColor Green
        Write-Host "  PFX (with private key): $((Resolve-Path $exportTo).Path)" -ForegroundColor Cyan
        Write-Host "  CER (public key only): $((Resolve-Path $cerPath).Path)" -ForegroundColor Cyan
        Write-Host "  Thumbprint: $thumbprint" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "The .cer file contains the public certificate that can be uploaded to Azure AD." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to export certificate: $_"
        Write-Host ""
       
    }
}
else {
    Write-Warning "Could not find certificate in store with thumbprint: $thumbprint"
    $certPath = "D:\Powershell\O365 Scripts\SiteProvisioning\$subject.cer"

    $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)

    [System.IO.File]::WriteAllBytes($certPath, $certBytes)
    Write-Host "PFX file created at: $exportTo" -ForegroundColor Yellow
    Write-Host ""
    
}