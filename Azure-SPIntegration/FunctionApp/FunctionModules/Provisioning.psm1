#$modulePath = Join-Path $PSScriptRoot "..\ExternalModules\PnP.PowerShell"

#$externalModules = Join-Path (Split-Path $PSScriptRoot -Parent) "ExternalModules"


#Import-Module Az.Accounts -Force -ErrorAction Stop
#Import-Module Az.Storage -Force -ErrorAction Stop

$script:ModuleRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($script:ModuleRoot)) {
    $script:ModuleRoot = Join-Path $PWD.Path "FunctionModules"
}
$script:AppRoot = Split-Path $script:ModuleRoot -Parent

$script:ContentTypeName = if ($env:PNP_CONTENT_TYPE_NAME) { $env:PNP_CONTENT_TYPE_NAME } else { 'content category page' }
$script:ContentTypeList = if ($env:PNP_CONTENT_TYPE_LIST) {
    $env:PNP_CONTENT_TYPE_LIST -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}
else {
    @('Site Pages')
}
$script:ViewsList = if ($env:PNP_VIEWS_LIST) {
    $env:PNP_VIEWS_LIST -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}
else {
    @('Documents', 'Site Pages')
}
$script:ViewFields = @(
    'DocIcon', 'Title', 'Modified', 'Editor', 'ReviewDate1',
    'DoogleWFMainCategory', 'MSDNotificationSent', 'DoogleWFRestrictedApproval', 'DoogleWFSubCategory'
)
$script:SiteColumnNames = @(
    'Main Category', 'Review Date', 'Notification Sent', 'Sub Category', 'Restricted Approval'
)

function Write-ProvisionLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info',

        [hashtable]$Properties = @{}
    )

    if (Get-Command -Name Write-ProvisionTrace -ErrorAction SilentlyContinue) {
        Write-ProvisionTrace -Message $Message -Level $Level -Properties $Properties
    }
    else {
        switch ($Level) {
            'Error' { Write-Host "ERROR: $Message" }
            'Warning' { Write-Warning $Message }
            default { Write-Host $Message }
        }
    }
}

function Invoke-ProvisionStep {
    param(
        [Parameter(Mandatory)]
        [string]$Step,

        [Parameter(Mandatory)]
        [scriptblock]$Action
    )

    Write-ProvisionLog "Starting $Step" -Properties @{ Step = $Step }

    try {
        $result = & $Action
        Write-ProvisionLog "Completed $Step" -Properties @{ Step = $Step }
        return $result
    }
    catch {
        Write-ProvisionLog "Failed $Step : $($_.Exception.Message)" -Level Error -Properties @{ Step = $Step }
        throw "[${Step}] $($_.Exception.Message)"
    }
}

function Get-AppRootPath {
    if (-not [string]::IsNullOrEmpty($script:AppRoot) -and (Test-Path $script:AppRoot)) {
        return $script:AppRoot
    }

    $runScriptRoot = $global:PSScriptRoot
    if (-not [string]::IsNullOrEmpty($runScriptRoot)) {
        return Split-Path $runScriptRoot -Parent
    }

    return $PWD.Path
}

function Test-AzureHostedEnvironment {
    return -not [string]::IsNullOrWhiteSpace($env:WEBSITE_SITE_NAME)
}

function Get-LoadedCertificateLocations {
    $roots = @('/var/ssl/private', 'C:\appservice\certificates\private')
    $locations = @()

    foreach ($root in $roots) {
        if (-not (Test-Path $root)) {
            continue
        }

        $files = @(Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue)
        foreach ($file in $files) {
            $locations += [PSCustomObject]@{
                Root       = $root
                FileName   = $file.Name
                FullPath   = $file.FullName
                Thumbprint = $file.BaseName.Replace(' ', '').ToUpperInvariant()
                Extension  = $file.Extension.ToLowerInvariant()
            }
        }
    }

    return $locations
}

function Get-CertificatePasswordSecureString {
    if ([string]::IsNullOrWhiteSpace($env:SPO_CERT_PASSWORD)) {
        return ConvertTo-SecureString '' -AsPlainText -Force
    }

    return ConvertTo-SecureString $env:SPO_CERT_PASSWORD -AsPlainText -Force
}

function Get-X509CertificateForConnect {
    param(
        [Parameter(Mandatory)]
        [string]$CertificatePath,

        [string]$KeyPath
    )

    $storageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
    $password = Get-CertificatePasswordSecureString

    if (-not [string]::IsNullOrWhiteSpace($KeyPath)) {
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($CertificatePath, $KeyPath)
    }

    try {
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $CertificatePath,
            $password,
            $storageFlags
        )
    }
    catch {
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $CertificatePath,
            '',
            $storageFlags
        )
    }
}

function Resolve-SharePointCertificate {
    param(
        [Parameter(Mandatory)]
        [string]$Thumbprint
    )

    $normalizedThumbprint = $Thumbprint.Replace(' ', '').ToUpperInvariant()
    $loadedCertificates = @(Get-LoadedCertificateLocations)
    $privateKeyCertificates = @($loadedCertificates | Where-Object { $_.Extension -in '.p12', '.pfx' })

    $matchedCertificate = @($privateKeyCertificates | Where-Object { $_.Thumbprint -eq $normalizedThumbprint } | Select-Object -First 1)[0]
    if ($null -ne $matchedCertificate) {
        return @{
            CertificatePath = $matchedCertificate.FullPath
        }
    }

    $keyFile = @($loadedCertificates | Where-Object {
            $_.Extension -eq '.key' -and $_.Thumbprint -eq $normalizedThumbprint
        } | Select-Object -First 1)[0]
    if ($null -ne $keyFile) {
        $crtPath = Join-Path '/var/ssl/certs' "$normalizedThumbprint.crt"
        if (Test-Path $crtPath) {
            return @{
                CertificatePath = $crtPath
                KeyPath         = $keyFile.FullPath
            }
        }
    }

    if ($privateKeyCertificates.Count -eq 1) {
        Write-ProvisionLog "SPO_CERT_THUMBPRINT '$normalizedThumbprint' did not match loaded certificate '$($privateKeyCertificates[0].Thumbprint)'. Using the only loaded certificate." -Level Warning
        return @{
            CertificatePath = $privateKeyCertificates[0].FullPath
        }
    }

    if (Test-AzureHostedEnvironment) {
        $availableCertificates = if ($loadedCertificates.Count -gt 0) {
            ($loadedCertificates | ForEach-Object { $_.FileName }) -join ', '
        }
        else {
            'none found under /var/ssl/private or C:\appservice\certificates\private'
        }

        throw "Certificate '$normalizedThumbprint' was not found in the Azure certificate store. Loaded certificates: $availableCertificates. Upload the .pfx under Certificates, set WEBSITE_LOAD_CERTIFICATES, update SPO_CERT_THUMBPRINT to the portal thumbprint, and restart the function app."
    }

    return @{
        Thumbprint = $normalizedThumbprint
    }
}

function Connect-PnPWithCertificate {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    if ([string]::IsNullOrWhiteSpace($env:SPO_CLIENT_ID)) {
        throw 'SPO_CLIENT_ID is not configured'
    }

    if ([string]::IsNullOrWhiteSpace($env:SPO_TENANT_ID)) {
        throw 'SPO_TENANT_ID is not configured'
    }

    $connectParams = @{
        Url         = $Url
        ClientId    = $env:SPO_CLIENT_ID
        Tenant      = $env:SPO_TENANT_ID
        ErrorAction = 'Stop'
    }

    $thumbprint = $env:SPO_CERT_THUMBPRINT
    if (-not [string]::IsNullOrWhiteSpace($thumbprint)) {
        $certAuth = Resolve-SharePointCertificate -Thumbprint $thumbprint

        if ($certAuth.ContainsKey('CertificatePath')) {
            Write-ProvisionLog "Connecting to SharePoint using certificate file $($certAuth.CertificatePath)"
            $certificate = Get-X509CertificateForConnect -CertificatePath $certAuth.CertificatePath -KeyPath $certAuth.KeyPath
            $certificateBase64 = [Convert]::ToBase64String(
                $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12)
            )
            Connect-PnPOnline @connectParams -CertificateBase64Encoded $certificateBase64
            return
        }

        if (-not (Test-AzureHostedEnvironment)) {
            Write-ProvisionLog "Connecting to SharePoint using certificate thumbprint $($certAuth.Thumbprint)"
            Connect-PnPOnline @connectParams -Thumbprint $certAuth.Thumbprint
            return
        }
    }

    $appRoot = Get-AppRootPath
    $certPath = Join-Path $appRoot 'cert.pfx'

    if (-not (Test-Path $certPath)) {
        throw "Certificate not configured. Set SPO_CERT_THUMBPRINT (and WEBSITE_LOAD_CERTIFICATES on Azure), or deploy cert.pfx for local development."
    }

    Write-ProvisionLog "Connecting to SharePoint using certificate file $certPath"
    $certificate = Get-X509CertificateForConnect -CertificatePath $certPath
    $certificateBase64 = [Convert]::ToBase64String(
        $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12)
    )
    Connect-PnPOnline @connectParams -CertificateBase64Encoded $certificateBase64
}

function Invoke-SharePointProvisioning {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,
        [string]$BlobUrl = $env:PNP_TEMPLATE_BLOB_URL
    )
    #Import-Module (Get-ChildItem "$externalModules\Az.Accounts" -Recurse -Filter "*.psd1" | Select-Object -First 1).FullName -Force
    #Import-Module (Get-ChildItem "$externalModules\Az.Storage" -Recurse -Filter "*.psd1" | Select-Object -First 1).FullName -Force
    #Import-Module (Get-ChildItem "$externalModules\PnP.PowerShell" -Recurse -Filter "*.psd1" | Select-Object -First 1).FullName -Force

    $baseFolder = $script:ModuleRoot

    # 2. Match your EXACT directory casing: 'ExternalModules'
    if (Test-Path (Join-Path $baseFolder "ExternalModules")) {
        $modulesFolder = Join-Path $baseFolder "ExternalModules"
    }
    else {
        $modulesFolder = Join-Path $baseFolder "..\ExternalModules"
    }

    # 3. Explicitly target the exact nested structure
    $azAccountsPath = Join-Path $modulesFolder "Az.Accounts/3.0.1/Az.Accounts.psd1"
    $pnpPath = Join-Path $modulesFolder "PnP.PowerShell/3.2.0/PnP.PowerShell.psd1" # Update version string to match your layout

    # 4. Import using forward slashes (Linux preferred)
    Import-Module $azAccountsPath -Force -ErrorAction Stop
    Import-Module $pnpPath -Force -ErrorAction Stop

    if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) {
        throw "Az.Accounts loaded but Connect-AzAccount not found"
    }

    if ([string]::IsNullOrWhiteSpace($BlobUrl)) {
        throw "PNP_TEMPLATE_BLOB_URL is required (full blob URL, with SAS if the blob is not public)"
    }

    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        throw "PnP.PowerShell loaded but Connect-PnPOnline not found"
    }

    Invoke-ProvisionStep -Step 'Connect-PnP' {
        Connect-PnPWithCertificate -Url $SiteUrl
    }

    $templatePath = Invoke-ProvisionStep -Step 'DownloadTemplate' {
        Get-AssetFromStorage -BlobUrl $BlobUrl
    }

    Invoke-ProvisionStep -Step 'Set-SiteRegionalSettings' {
        Set-SiteRegionalSettings -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Set-SearchSettings' {
        Set-SearchSettings -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Install-App' {
        Install-App -SiteUrl $SiteUrl
    }

    if ([string]::IsNullOrWhiteSpace($templatePath) -or -not (Test-Path $templatePath)) {
        throw "[ApplyTemplate] Template file not found at $templatePath"
    }

    Invoke-ProvisionStep -Step 'Invoke-PnPSiteTemplate' {
        Write-ProvisionLog "Applying template: $templatePath" -Properties @{ Step = 'Invoke-PnPSiteTemplate' }
        Invoke-PnPSiteTemplate -Path $templatePath
    }
    Invoke-ProvisionStep -Step 'Add-GroupstoSharePointGroups' {
        Add-GroupstoSharePointGroups -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Set-DocLibraryPermissions' {
        Set-DocLibraryPermissions -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Set-Branding' {
        Set-Branding -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Add-ContentTypes' {
        Add-ContentTypes -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Add-SiteColumns' {
        Add-SiteColumns -SiteUrl $SiteUrl
    }
    Invoke-ProvisionStep -Step 'Set-Views' {
        Set-Views -SiteUrl $SiteUrl
    }
    if (-not [string]::IsNullOrWhiteSpace($env:SPO_HUB_SITE_URL)) {
        Invoke-ProvisionStep -Step 'Add-HubSites' {
            Add-HubSites -SiteUrl $SiteUrl
        }
    }
    else {
        Write-ProvisionLog 'Skipping hub site association; SPO_HUB_SITE_URL is not configured' -Properties @{ Step = 'Add-HubSites' }
    }

    $webTitle = $SiteUrl.Split('/')[-1]
    try {
        $summaryWeb = Get-PnPWeb -Includes Title
        Invoke-PnPQuery
        if (-not [string]::IsNullOrWhiteSpace($summaryWeb.Title)) {
            $webTitle = $summaryWeb.Title
        }
    }
    catch {
        Write-ProvisionLog "Could not read site title; using site name '$webTitle'" -Properties @{ Step = 'Complete' }
    }

    return @{
        status   = 'Success'
        message  = 'SharePoint provisioning completed'
        siteUrl  = $SiteUrl
        webTitle = $webTitle
    }
}


function Set-SiteRegionalSettings {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [string]$SiteName
    )

    try {
        Write-ProvisionLog "Updating regional settings for $SiteUrl" -Properties @{ Step = 'Set-SiteRegionalSettings' }

        $web = Get-PnPWeb -Includes RegionalSettings
        $localeId = 5129
        $web.RegionalSettings.LocaleId = $localeId
        $web.Update()
        Invoke-PnPQuery

        Write-ProvisionLog "Updated regional settings for $SiteUrl" -Properties @{ Step = 'Set-SiteRegionalSettings' }

        return @{
            status   = 'Success'
            message  = 'Updated Site Regional Settings to New Zealand locale'
            siteUrl  = $SiteUrl
            localeId = $localeId
        }
    }
    catch {
        Write-ProvisionLog "Error updating regional settings for $SiteUrl : $($_.Exception.Message)" -Level Error -Properties @{ Step = 'Set-SiteRegionalSettings' }
        throw
    }
}


function Set-SearchSettings {
    param (
        [string]$SiteUrl
    )
    try {

        $list = Get-PnPList -Identity "Site Assets" -ErrorAction SilentlyContinue
        if ($list) {
            $list.NoCrawl = $true
            $list.Update()
            Invoke-PnPQuery
            Write-ProvisionLog "Site Assets list no crawled on $($SiteUrl)" 
        }
        else {
            $web = Get-PnPWeb -Includes Lists
            Invoke-PnPQuery
            $web.Lists.EnsureSiteAssetsLibrary()
            Invoke-PnPQuery
            $list = Get-PnPList -Identity "Site Assets" 
            Set-PnPList -Identity $list -NoCrawl:$true
            Invoke-PnPQuery
            Write-ProvisionLog "Site Assets list no crawled on $($SiteUrl)"
        }

    }
    catch {
        Write-ProvisionLog "ERROR: Failed to set search settings on $($SiteUrl): $($_.Exception.Message)"
       
    }
}

function Install-App {
    param(
        [string]$SiteUrl
    )
    try {
        #$appIds = @('59903278-DD5D-4E9E-BEF6-562AAE716B8B', '00406271-0276-406F-9666-512623EB6709')
        $appIdsString = $env:PNP_SPFX_APP_IDS
        if ([string]::IsNullOrWhiteSpace($appIdsString)) {
            throw "PNP_SPFX_APP_IDS is not configured"
        }
        $appIds = $appIdsString.Split(',')
        foreach ($appId in $appIds) {
            $app = Get-PnPApp -Identity $appId -ErrorAction Stop
            $appTitle = if (-not [string]::IsNullOrWhiteSpace($app.Title)) { $app.Title } else { $appId }

            if ($null -eq $app.InstalledVersion) {
                Install-PnPApp -Identity $app -ErrorAction Stop
                Write-ProvisionLog "App installed: $appTitle" -Properties @{ Step = 'Install-App'; AppId = $appId }
            }
            else {
                Write-ProvisionLog "App already installed: $appTitle (version $($app.InstalledVersion))" -Properties @{ Step = 'Install-App'; AppId = $appId }
            }
        }
    }
    catch {
        Write-ProvisionLog "Failed to install app on $SiteUrl : $($_.Exception.Message)" -Level Error -Properties @{ Step = 'Install-App' }
        throw
    }
}


function Add-GroupstoSharePointGroups {
    param (
        [string]$SiteUrl
    )

    try {
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop

        #$EntraOwnerGroupObjectIds = @('b4fa1e98-2893-4f26-bc64-f7a3e93b3753', '4a6457ed-7d02-4faf-9a9e-c834e7dbed20')
        $EntraOwnerGroupObjectIdsString = $env:PNP_ENTRA_OWNER_OBJECT_IDS
        if ([string]::IsNullOrWhiteSpace($EntraOwnerGroupObjectIdsString)) {
            throw "PNP_ENTRA_OWNER_OBJECT_IDS is not configured"
        }
        $EntraOwnerGroupObjectIds = $EntraOwnerGroupObjectIdsString.Split(',')
        foreach ($EntraGroupObjectId in $EntraOwnerGroupObjectIds) {
            $groupLoginName = "c:0t.c|tenant|$EntraGroupObjectId"
            New-PnPUser -LoginName $groupLoginName -ErrorAction SilentlyContinue | Out-Null
            Add-PnPGroupMember -Group $ownersGroup -LoginName $groupLoginName -ErrorAction Stop
            Write-ProvisionLog "Added Entra group $EntraGroupObjectId to Owners group on $SiteUrl"
        }

        #$EntraMemberGroupObjectIds = @('4a6457ed-7d02-4faf-9a9e-c834e7dbed20')
        $EntraMemberGroupObjectIdsString = $env:PNP_ENTRA_MEMBER_OBJECT_IDS
        if ([string]::IsNullOrWhiteSpace($EntraMemberGroupObjectIdsString)) {
            throw "PNP_ENTRA_MEMBER_OBJECT_IDS is not configured"
        }
        $EntraMemberGroupObjectIds = $EntraMemberGroupObjectIdsString.Split(',')
        foreach ($EntraGroupObjectId in $EntraMemberGroupObjectIds) {
            $groupLoginName = "c:0t.c|tenant|$EntraGroupObjectId"
            New-PnPUser -LoginName $groupLoginName -ErrorAction SilentlyContinue | Out-Null
            Add-PnPGroupMember -Group $membersGroup -LoginName $groupLoginName -ErrorAction Stop
            Write-ProvisionLog "Added Entra group $EntraGroupObjectId to Members group on $SiteUrl"
        }
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to add group to SharePoint group on $($SiteUrl): $($_.Exception.Message)"
    }
}

function Set-DocLibraryPermissions {
    param(
        [string]$SiteUrl
    )

    try {
        Write-ProvisionLog "Setting Documents library permissions on $SiteUrl" -Properties @{ Step = 'Set-DocLibraryPermissions' }

        $library = Get-PnPList -Identity 'Documents' -ErrorAction Stop
        Set-PnPList -Identity $library -BreakRoleInheritance -CopyRoleAssignments -ErrorAction Stop

        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        $membersGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop

        if (-not $ownersGroup -or -not $membersGroup) {
            Write-ProvisionLog "ERROR: Could not resolve Owners or Members group for $SiteUrl" -Properties @{ Step = 'Set-DocLibraryPermissions' }
            return
        }

        Set-PnPListPermission `
            -Identity $library `
            -Group $ownersGroup `
            -RemoveRole 'Full Control' `
            -AddRole 'Contribute' `
            -ErrorAction Stop

        Set-PnPListPermission `
            -Identity $library `
            -Group $membersGroup `
            -RemoveRole 'Edit' `
            -AddRole 'Contribute' `
            -ErrorAction Stop

        Write-ProvisionLog 'Permissions updated: Owners and Members have Contribute on Documents' -Properties @{ Step = 'Set-DocLibraryPermissions' }
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to set DocLibraryPermissions on $($SiteUrl): $($_.Exception.Message)" -Properties @{ Step = 'Set-DocLibraryPermissions' }
    }
}



function Get-StorageAccessToken {
    Connect-AzAccount -Identity -ErrorAction Stop | Out-Null

    $token = (Get-AzAccessToken -ResourceUrl 'https://storage.azure.com' -ErrorAction Stop).Token
    if ($token -is [SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($token)
        try {
            return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }

    return [string]$token
}

function Get-AssetFromStorage {
    param(
        [Parameter(Mandatory)]
        [string]$BlobUrl
    )

    $uri = [Uri]$BlobUrl
    $fileName = [System.IO.Path]::GetFileName($uri.LocalPath)

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        throw "Could not determine file name from blob URL: $BlobUrl"
    }

    $tempRoot = if ([string]::IsNullOrWhiteSpace($env:TEMP)) { '/tmp' } else { $env:TEMP }
    $tempFolder = Join-Path $tempRoot "PnPTemplates"

    if (-not (Test-Path $tempFolder)) {
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    }

    $localPath = Join-Path $tempFolder $fileName
    $downloadUrl = $BlobUrl

    if ($uri.Query -notmatch '(^|[?&])sig=' -and -not [string]::IsNullOrWhiteSpace($env:PNP_STORAGE_BLOB_SAS)) {
        $sas = $env:PNP_STORAGE_BLOB_SAS.Trim()
        if (-not $sas.StartsWith('?')) {
            $sas = "?$sas"
        }
        $downloadUrl = "$BlobUrl$sas"
        $uri = [Uri]$downloadUrl
    }

    Write-ProvisionLog "Downloading blob from $downloadUrl"

    try {
        if ($uri.Query -match '(^|[?&])sig=') {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $localPath -UseBasicParsing -ErrorAction Stop
        }
        else {
            Write-ProvisionLog 'Authenticating to Azure Storage with managed identity' -Properties @{ Step = 'DownloadTemplate' }

            $accessToken = Get-StorageAccessToken
            $headers = @{
                Authorization  = "Bearer $accessToken"
                'x-ms-version' = '2021-08-06'
            }

            Invoke-WebRequest -Uri $downloadUrl -Headers $headers -OutFile $localPath -UseBasicParsing -ErrorAction Stop
        }
    }
    catch {
        $statusCode = $null
        if ($null -ne $_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        elseif ($null -ne $_.Exception.StatusCode) {
            $statusCode = [int]$_.Exception.StatusCode
        }

        if ($statusCode -eq 401) {
            throw "Azure Storage returned 401 for '$BlobUrl'. Ensure the Function App managed identity has 'Storage Blob Data Reader' (or Contributor) on the storage account, or set PNP_STORAGE_BLOB_SAS / include a SAS token in the blob URL."
        }

        throw
    }

    if (-not (Test-Path $localPath)) {
        throw "Blob download failed; file not found at $localPath"
    }

    return $localPath
}


function Resolve-UploadedFileServerRelativeUrl {
    param(
        [Parameter(Mandatory)]
        [string]$LocalFilePath,

        [Parameter(Mandatory)]
        [string]$FolderName,

        $UploadedFile
    )

    $fileName = [System.IO.Path]::GetFileName($LocalFilePath)

    if ($null -ne $UploadedFile) {
        if (-not [string]::IsNullOrWhiteSpace($UploadedFile.ServerRelativeUrl)) {
            return $UploadedFile.ServerRelativeUrl
        }

        try {
            $loadedUrl = (Get-PnPProperty -ClientObject $UploadedFile -Property ServerRelativeUrl -ErrorAction Stop).ServerRelativeUrl
            if (-not [string]::IsNullOrWhiteSpace($loadedUrl)) {
                return $loadedUrl
            }
        }
        catch {
            Write-ProvisionLog "Could not load ServerRelativeUrl from upload result; resolving path from web" -Properties @{
                Step = 'Set-Branding'
            }
        }
    }

    $web = Get-PnPWeb -Includes ServerRelativeUrl
    Invoke-PnPQuery
    $baseUrl = $web.ServerRelativeUrl.TrimEnd('/')

    $candidatePaths = @(
        "$baseUrl/$FolderName/$fileName"
        "$baseUrl/SiteAssets/$fileName"
        "$baseUrl/siteassets/$fileName"
    )

    foreach ($candidatePath in $candidatePaths) {
        try {
            $file = Get-PnPFile -Url $candidatePath -AsFileObject -ErrorAction Stop
            $resolvedUrl = (Get-PnPProperty -ClientObject $file -Property ServerRelativeUrl -ErrorAction Stop).ServerRelativeUrl
            if (-not [string]::IsNullOrWhiteSpace($resolvedUrl)) {
                return $resolvedUrl
            }
        }
        catch {
            continue
        }
    }

    return "$baseUrl/SiteAssets/$fileName"
}

function Set-Branding {
    param (
        [string]$SiteUrl,
        [string]$BlobUrl = $env:PNP_BRANDING_BLOB_URL,
        [string]$HeadeBlobUrl = $env:PNP_HEADER_BLOB_URL
    )

    if ([string]::IsNullOrWhiteSpace($BlobUrl)) {
        throw 'PNP_BRANDING_BLOB_URL is required (full blob URL, with SAS if the blob is not public)'
    }
    if ([string]::IsNullOrWhiteSpace($HeadeBlobUrl)) {
        throw 'PNP_HEADER_BLOB_URL is required (full blob URL, with SAS if the blob is not public)'
    }

    Write-ProvisionLog 'Setting branding' -Properties @{ Step = 'Set-Branding' }

    $web = Get-PnPWeb -Includes Lists
    Invoke-PnPQuery
    $web.Lists.EnsureSiteAssetsLibrary()
    Invoke-PnPQuery

    Set-PnPWebHeader -HeaderLayout Extended -ErrorAction Stop

    $imagePath = Get-AssetFromStorage -BlobUrl $BlobUrl
    $headerImagePath = Get-AssetFromStorage -BlobUrl $HeadeBlobUrl  
    $headerFileName = [System.IO.Path]::GetFileName($headerImagePath)
    $spFile = Add-PnPFile -Path $headerImagePath -Folder 'SiteAssets' -NewFileName $headerFileName -ErrorAction Stop
    $headerFileUrl = Resolve-UploadedFileServerRelativeUrl -LocalFilePath $headerImagePath -FolderName 'SiteAssets' -UploadedFile $spFile
    $fileName = [System.IO.Path]::GetFileName($imagePath)
    $spFile = Add-PnPFile -Path $imagePath -Folder 'SiteAssets' -NewFileName $fileName -ErrorAction Stop
    $fileUrl = Resolve-UploadedFileServerRelativeUrl -LocalFilePath $imagePath -FolderName 'SiteAssets' -UploadedFile $spFile

    if ([string]::IsNullOrWhiteSpace($fileUrl)) {
        throw "Uploaded file URL could not be resolved for '$fileName'"
    }

    $backgroundUrl = [System.Uri]::EscapeUriString($fileUrl)
    if ([string]::IsNullOrWhiteSpace($backgroundUrl)) {
        $backgroundUrl = $fileUrl
    }

    Write-ProvisionLog "Uploaded branding image to $fileUrl" -Properties @{
        Step            = 'Set-Branding'
        BackgroundImage = $backgroundUrl
        FileName        = $fileName
    }

    Set-PnPWebHeader `
        -HeaderLayout Extended `
        -HeaderBackgroundImageUrl $headerFileUrl `
        -HeaderBackgroundImageFocalX 0.5 `
        -HeaderBackgroundImageFocalY 0.5 `
        -ErrorAction Stop

    Set-PnPWebHeader `
        -HeaderBackgroundImageUrl $headerFileUrl `
        -HeaderLayout Extended `
        -ErrorAction Stop

    Set-PnPWebHeader -HeaderLayout Extended -ErrorAction Stop
    Set-PnPFooter -Layout Extended -ErrorAction Stop
    Set-PnPWeb -MegaMenuEnabled:$false -ErrorAction Stop
    Set-PnPWebHeader -SiteLogoUrl $fileUrl -ErrorAction Stop
    Set-PnPWebHeader -SiteThumbnailUrl $fileUrl -ErrorAction Stop

    Write-ProvisionLog 'Branding applied' -Properties @{
        Step            = 'Set-Branding'
        BackgroundImage = $backgroundUrl
    }

    return @{
        status          = 'Success'
        message         = 'Branding applied'
        siteUrl         = $SiteUrl
        backgroundImage = $backgroundUrl
    }
}

function Get-ContentTypeHub {
    param(
        [string]$SiteUrl,
        [string]$ContentTypeNames = $script:ContentTypeName
    )

    try {
        Write-ProvisionLog "Adding content types from Content Type Hub for $SiteUrl"

        $contentTypesArray = $ContentTypeNames.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl
        Write-ProvisionLog "Content Type Hub URL: $contentTypeHubUrl"

        Connect-PnPWithCertificate -Url $contentTypeHubUrl
        $contentTypesToAdd = @(Get-PnPContentType |
            Where-Object { $contentTypesArray -contains $_.Name } |
            ForEach-Object {
                [PSCustomObject]@{
                    Name     = $_.Name
                    StringId = if ($_.StringId) { $_.StringId } else { $_.Id.StringValue }
                }
            })

        Connect-PnPWithCertificate -Url $SiteUrl

        foreach ($contentType in $contentTypesToAdd) {
            Add-PnPContentTypesFromContentTypeHub -ContentTypes $contentType.StringId -ErrorAction Stop
            Write-ProvisionLog "Added content type '$($contentType.Name)' from hub to site: $SiteUrl"
        }
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to add content types from hub on $($SiteUrl): $($_.Exception.Message)"
    }
}


function Add-ContentTypes {
    param(
        [string]$SiteUrl,
        [string]$ContentTypeName = $script:ContentTypeName,
        [string[]]$TargetLists = $script:ContentTypeList
    )

    try {
        Write-ProvisionLog "Adding content types on $SiteUrl"

        $library = Get-PnPList -Identity 'Documents' -ErrorAction Stop
        if (-not $library) {
            Write-ProvisionLog "ERROR: Documents library not found on $SiteUrl"
            return
        }

        $library.ContentTypesEnabled = $true
        $library.Update()
        Invoke-PnPQuery
        Write-ProvisionLog "Content types enabled on $($library.Title)"

        Get-ContentTypeHub -SiteUrl $SiteUrl -ContentTypeNames $ContentTypeName

        $contentType = Get-PnPContentType -Identity $ContentTypeName -ErrorAction SilentlyContinue
        if (-not $contentType) {
            Write-ProvisionLog "ERROR: Content type '$ContentTypeName' not found on $SiteUrl"
            return
        }

        foreach ($listName in $TargetLists) {
            Add-PnPContentTypeToList -List $listName -ContentType $contentType -ErrorAction Stop
            Write-ProvisionLog "Added content type '$($contentType.Name)' to list '$listName'"
            Set-DefaultContentType -SiteUrl $SiteUrl -ListName $listName -ContentTypeName $ContentTypeName
        }
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to add content type on $($SiteUrl): $($_.Exception.Message)"
    }
}

function Set-DefaultContentType {
    param(
        [string]$SiteUrl,
        [string]$ListName,
        [string]$ContentTypeName
    )

    try {
        Set-PnPDefaultContentTypeToList -List $ListName -ContentType $ContentTypeName -ErrorAction Stop
        Write-ProvisionLog "Default content type set to '$ContentTypeName' on list '$ListName' ($SiteUrl)"
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to set default content type '$ContentTypeName' on list '$ListName' ($SiteUrl): $($_.Exception.Message)"
    }
}


function Add-SiteColumnToList {
    param(
        [Parameter(Mandatory)]
        $List,

        [Parameter(Mandatory)]
        [string]$ListName,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        $ExistingColumn,

        [Parameter(Mandatory)]
        [string]$ColumnName
    )

    $fieldInList = Get-PnPField -List $List -Identity $ColumnName -ErrorAction SilentlyContinue
    if ($fieldInList) {
        Write-ProvisionLog "Site column '$ColumnName' already exists in '$ListName' on $SiteUrl. Skipping."
        return
    }

    $schemaXml = (Get-PnPProperty -ClientObject $ExistingColumn -Property SchemaXml).SchemaXml

    switch ($ColumnName) {
        'Main Category' {
            Add-PnPFieldFromXml -List $List -FieldXml $schemaXml -ErrorAction Stop
        }
        'Sub Category' {
            Add-PnPFieldFromXml -List $List -FieldXml $schemaXml -ErrorAction Stop
        }
        Default {
            Add-PnPField -List $List -Field $ExistingColumn -ErrorAction Stop
        }
    }

    Write-ProvisionLog "Added site column '$ColumnName' to '$ListName' on $SiteUrl"
}

function Add-SiteColumns {
    param (
        [string]$SiteUrl,
        [string[]]$TargetLists = @('Documents', 'Site Pages')
    )

    try {
        Write-ProvisionLog "Adding site columns on $SiteUrl"

        foreach ($columnName in $script:SiteColumnNames) {
            $existingColumn = Get-PnPField -Identity $columnName -Includes SchemaXml -ErrorAction SilentlyContinue
            if (-not $existingColumn) {
                Write-ProvisionLog "Site column '$columnName' not found on $SiteUrl. Skipping."
                continue
            }

            Write-ProvisionLog "Site column '$columnName' exists on $SiteUrl"

            if ($columnName -eq 'Review Date') {
                $existingColumn.DefaultFormula = '=TODAY()+365'
                $existingColumn.UpdateAndPushChanges($true)
                Invoke-PnPQuery
            }

            foreach ($listName in $TargetLists) {
                $list = Get-PnPList -Identity $listName -ErrorAction Stop
                Add-SiteColumnToList -List $list -ListName $listName -SiteUrl $SiteUrl -ExistingColumn $existingColumn -ColumnName $columnName
            }
        }
    }
    catch {
        Write-ProvisionLog "ERROR: Failed to add site column on $($SiteUrl): $($_.Exception.Message)" -Level Error
        throw
    }
}

function Set-DocumentsViews {
    param(
        [Parameter(Mandatory)]
        $List
    )

    $viewFields = $script:ViewFields
    $viewsToUpdate = @(
        Get-PnPView -List $List | Where-Object { $_.DefaultView -eq $true }
    )

    if ($viewsToUpdate.Count -lt 1) {
        $viewsToUpdate = @(Get-PnPView -List $List | Where-Object { $_.Title -eq 'All Documents' })
    }

    $updatedViews = 0
    foreach ($view in $viewsToUpdate) {
        Write-ProvisionLog "Updating Documents view '$($view.Title)'" -Properties @{
            Step   = 'Set-Views'
            View   = $view.Title
            Fields = ($viewFields -join ', ')
        }

        try {
            Set-PnPView -List $List -Identity $view.Title -Fields $viewFields -ErrorAction Stop
            $updatedViews++
            Write-ProvisionLog "Updated Documents view '$($view.Title)'" -Properties @{ Step = 'Set-Views' }
        }
        catch {
            Write-ProvisionLog "Skipping Documents view '$($view.Title)': $($_.Exception.Message)" -Level Warning -Properties @{ Step = 'Set-Views' }
        }
    }

    if ($updatedViews -lt 1) {
        throw 'No Documents views could be updated.'
    }
}

function Set-SitePagesViews {
    param(
        [Parameter(Mandatory)]
        $List
    )

    # Hardcoded Site Pages views and SharePoint internal field names.
    $viewNames = @(
        'All Pages',
        'By Editor',
        'By Author',
        'Created By Me'
    )

    $viewFields = @(
        'DocIcon',
        'LinkFilename',
        'Editor',
        'Modified',
        'Author',
        'Created',
        'ReviewDate1',
        'DoogleWFMainCategory',
        'MSDNotificationSent',
        'DoogleWFRestrictedApproval',
        'DoogleWFSubCategory'
    )

    $viewsToUpdate = @(
        Get-PnPView -List $List | Where-Object { $_.DefaultView -eq $true }
    )

    if ($viewsToUpdate.Count -lt 1) {
        $viewsToUpdate = @($viewNames | ForEach-Object {
                Get-PnPView -List $List -Identity $_ -ErrorAction SilentlyContinue
            } | Where-Object { $null -ne $_ })
    }

    $updatedViews = 0
    foreach ($view in $viewsToUpdate) {
        Write-ProvisionLog "Updating Site Pages view '$($view.Title)'" -Properties @{
            Step   = 'Set-Views'
            View   = $view.Title
            Fields = ($viewFields -join ', ')
        }

        try {
            Set-PnPView -List $List -Identity $view.Title -Fields $viewFields -ErrorAction Stop
            $updatedViews++
            Write-ProvisionLog "Updated Site Pages view '$($view.Title)'" -Properties @{ Step = 'Set-Views' }
        }
        catch {
            Write-ProvisionLog "Skipping Site Pages view '$($view.Title)': $($_.Exception.Message)" -Level Warning -Properties @{ Step = 'Set-Views' }
        }
    }

    if ($updatedViews -lt 1) {
        throw 'No Site Pages views could be updated.'
    }
}

function Get-ViewFieldsForList {
    param(
        [Parameter(Mandatory)]
        [string]$ListName
    )

    return $script:ViewFields
}

function Resolve-ViewFieldName {
    param(
        [Parameter(Mandatory)]
        [string]$Field,

        [Parameter(Mandatory)]
        [string[]]$ListFieldNames
    )

    if ($ListFieldNames -contains $Field) {
        return $Field
    }

    switch ($Field) {
        'Name' {
            if ($ListFieldNames -contains 'LinkTitle') {
                return 'LinkTitle'
            }
            if ($ListFieldNames -contains 'LinkFilename') {
                return 'LinkFilename'
            }
            if ($ListFieldNames -contains 'FileLeafRef') {
                return 'FileLeafRef'
            }
        }
        'Editor' {
            if ($ListFieldNames -contains 'Editor') {
                return 'Editor'
            }
        }
        'Author' {
            if ($ListFieldNames -contains 'Author') {
                return 'Author'
            }
        }
        'PromotedState' {
            foreach ($candidate in @('PromotedState', '_ModerationStatus')) {
                if ($ListFieldNames -contains $candidate) {
                    return $candidate
                }
            }
        }
        'Title' {
            if ($ListFieldNames -contains 'LinkTitle') {
                return 'LinkTitle'
            }
            if ($ListFieldNames -contains 'LinkFilename') {
                return 'LinkFilename'
            }
        }
    }

    return $null
}

function Resolve-AvailableViewFields {
    param(
        [Parameter(Mandatory)]
        [string[]]$DesiredFields,

        [Parameter(Mandatory)]
        [string[]]$ListFieldNames
    )

    $availableFields = @()
    foreach ($field in $DesiredFields) {
        $resolvedField = Resolve-ViewFieldName -Field $field -ListFieldNames $ListFieldNames
        if (-not [string]::IsNullOrWhiteSpace($resolvedField)) {
            $availableFields += $resolvedField
        }
    }

    return $availableFields
}

function Get-ListViewIdentitiesToUpdate {
    param(
        [Parameter(Mandatory)]
        [string]$ListName,

        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    $views = @(Get-PnPView -List $ListName -ErrorAction Stop)

    return @(Get-DefaultListViewIdentity -ListName $ListName -SiteUrl $SiteUrl -Views $views)
}

function Get-DefaultListViewIdentity {
    param(
        [Parameter(Mandatory)]
        [string]$ListName,

        [Parameter(Mandatory)]
        [string]$SiteUrl,

        $Views
    )

    if ($null -eq $Views) {
        $Views = @(Get-PnPView -List $ListName -ErrorAction Stop)
    }

    $defaultView = @($Views | Where-Object { $_.DefaultView -eq $true } | Select-Object -First 1)[0]

    if ($null -ne $defaultView -and -not [string]::IsNullOrWhiteSpace($defaultView.Title)) {
        return $defaultView.Title
    }

    $fallbackViewNames = @{
        'Documents'  = @('All Documents')
        'Site Pages' = @('All Pages', 'By Editor', 'Recent')
    }

    if ($fallbackViewNames.ContainsKey($ListName)) {
        foreach ($candidate in $fallbackViewNames[$ListName]) {
            if (@($Views | Where-Object { $_.Title -eq $candidate }).Count -gt 0) {
                return $candidate
            }
        }

        return $fallbackViewNames[$ListName][0]
    }

    if ($Views.Count -gt 0) {
        return $Views[0].Title
    }

    throw "Default view identity could not be resolved for list '$ListName' ($SiteUrl)"
}

function Set-Views {
    param(
        [string]$SiteUrl,
        [string[]]$Lists = $script:ViewsList,
        [string[]]$Fields = $script:ViewFields
    )

    Write-ProvisionLog 'Setting views' -Properties @{ Step = 'Set-Views' }

    foreach ($listName in $Lists) {
        $library = Get-PnPList -Identity $listName -ErrorAction Stop
        if (-not $library) {
            throw "List '$listName' not found on $SiteUrl"
        }

        if ($listName -eq 'Site Pages') {
            Set-SitePagesViews -List $library
            continue
        }

        if ($listName -eq 'Documents') {
            Set-DocumentsViews -List $library
            continue
        }

        $viewIdentities = @(Get-ListViewIdentitiesToUpdate -ListName $listName -SiteUrl $SiteUrl)
        $desiredFields = Get-ViewFieldsForList -ListName $listName

        $listFieldNames = @(Get-PnPField -List $library -ErrorAction Stop | ForEach-Object { $_.InternalName })
        $availableFields = Resolve-AvailableViewFields -DesiredFields $desiredFields -ListFieldNames $listFieldNames

        if ($availableFields.Count -lt 1) {
            throw "No configured view fields exist on list '$listName'. Desired fields: $($desiredFields -join ', ')"
        }

        foreach ($viewIdentity in $viewIdentities) {
            Write-ProvisionLog "Updating view '$viewIdentity' on '$listName'" -Properties @{
                Step   = 'Set-Views'
                View   = [string]$viewIdentity
                Fields = ($availableFields -join ', ')
            }

            Set-PnPView -List $library -Identity $viewIdentity -Fields $availableFields -ErrorAction Stop
            Write-ProvisionLog "Updated view '$viewIdentity' on '$listName'" -Properties @{ Step = 'Set-Views' }
        }
    }
}


function Add-HubSites {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    try {
        $hubSite = $env:SPO_HUB_SITE_URL
        if ([string]::IsNullOrWhiteSpace($hubSite)) {
            throw 'SPO_HUB_SITE_URL is not configured'
        }

        Write-ProvisionLog "Associating $SiteUrl with hub site $hubSite" -Properties @{ Step = 'Add-HubSites' }

        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $hubSite -ErrorAction Stop
        Write-ProvisionLog "Hub site associated with $SiteUrl" -Properties @{ Step = 'Add-HubSites' }

        return @{
            status  = 'Success'
            message = "Hub site associated with $SiteUrl"
            siteUrl = $SiteUrl
            hubSite = $hubSite
        }
    }
    catch {
        Write-ProvisionLog "Failed to associate hub site on $SiteUrl : $($_.Exception.Message)" -Level Error -Properties @{ Step = 'Add-HubSites' }
        throw
    }

}



Export-ModuleMember -Function Invoke-SharePointProvisioning, Set-SiteRegionalSettings, Set-SearchSettings, Install-App, Add-GroupstoSharePointGroups, Set-DocLibraryPermissions, Set-Branding, Add-ContentTypes, Add-SiteColumns, Set-Views, Add-HubSites