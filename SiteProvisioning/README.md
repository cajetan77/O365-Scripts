
# Steps – App Registration with Certificate-Based Authentication

## Objective

Create an App Registration in Microsoft Entra ID and configure certificate-based authentication for access to both **Microsoft Graph API** and **SharePoint API**.

---

## Implementation Steps

### 1. Create App Registration

1. Log in to **Azure Portal**
2. Navigate to **Microsoft Entra ID**
3. Select **App registrations**
4. Click **New registration**
5. Enter the application name
6. Select the required supported account type
7. Click **Register**

---

### 2. Capture Application Details

Record the following:

* **Application (Client) ID**
* **Directory (Tenant) ID**

---

### 3. Upload Certificate

1. Open the App Registration
2. Navigate to **Certificates & secrets**
3. Select **Certificates**
4. Click **Upload certificate**
5. Upload the **.cer** file
6. Confirm the certificate is added successfully

---

### 4. Add Microsoft Graph API Permission

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Add:

   * **Sites.FullControl.All**
6. Click **Add permissions**

---

### 5. Add SharePoint API Permission

1. In **API permissions**, click **Add a permission**
2. Select **SharePoint**
3. Choose **Application permissions**
4. Add:

   * **Sites.FullControl.All**
5. Click **Add permissions**

---

### 6. Grant Admin Consent

1. In **API permissions**, click **Grant admin consent**
2. Confirm consent for both:

   * **Microsoft Graph – Sites.FullControl.All**
   * **SharePoint – Sites.FullControl.All**

---

### 7. Configure Application Authentication

1. Store the certificate private key securely
2. Configure the application with:

   * Tenant ID
   * Client ID
   * Certificate thumbprint or certificate reference

---

### 8. Validate Access

1. Request an access token using the certificate
2. Test **Microsoft Graph API** access to SharePoint sites
3. Test **SharePoint API** access
4. Confirm the application can authenticate and access SharePoint successfully

---

## Validation

* App Registration created successfully
* Certificate uploaded successfully
* **Microsoft Graph – Sites.FullControl.All** added
* **SharePoint – Sites.FullControl.All** added
* Admin consent granted
* Authentication using certificate is successful
* Access to both Graph and SharePoint APIs is successful

---

## Rollback Plan

* Remove Microsoft Graph API permission
* Remove SharePoint API permission
* Remove uploaded certificate
* Delete App Registration if no longer required

---


# Set-SharePointSitePropertiesFromCsv.ps1

This script bulk-updates SharePoint Online sites listed in a CSV file by applying common site settings and governance controls.

## What the script does

For each `SiteUrl` in the CSV, the script connects using app-only authentication and performs the following:

- Sets regional settings (locale/timezone related settings in the site web).
- Updates search settings (sets **Site Assets** library to no-crawl).
- Updates permissions on the **Documents** library (breaks inheritance and sets Owners/Members to Contribute).
- Adds a configured Entra ID group to the site Owners and Members SharePoint groups.
- Applies branding:
  - Header/footer layout changes
  - Uploads and applies site logo and thumbnail
- Installs configured SharePoint apps by App ID.
- Adds/publishes content types from the content type hub and sets default content type on target lists.

At the end, it prints success/failed totals for all processed rows.

## Pre-requisites

Before running, ensure the following are in place:

- PowerShell 7+ (recommended).
- `PnP.PowerShell` module installed:
  - `Install-Module PnP.PowerShell -Scope CurrentUser`
- SharePoint app registration with permissions required to:
  - Connect to target sites
  - Manage lists/libraries, groups, content types, branding, and apps
- Certificate-based auth details available for the app registration.
- A valid `config.json` file in the same folder (or pass `-ConfigPath`).
- A valid `Sites.csv` file in the same folder (or pass `-CsvPath`) with a `SiteUrl` column.
- Local files available for logo/thumbnail paths if branding is enabled.

## Required input files

### 1) `config.json`

Minimum required keys used by the script:

- `TenantId`
- `AppId`
- `TenantName`
- `ThumbPrint`
- `ClientSecret` (used when connecting to the content type hub in current script logic)

Example:

```json
{
  "TenantId": "00000000-0000-0000-0000-000000000000",
  "AppId": "11111111-1111-1111-1111-111111111111",
  "TenantName": "contoso.onmicrosoft.com",
  "ThumbPrint": "ABCDEF1234567890ABCDEF1234567890ABCDEF12",
  "ClientSecret": "your-client-secret"
}
```

### 2) `Sites.csv`

The CSV must include a `SiteUrl` header.

Example:

```csv
SiteUrl
https://contoso.sharepoint.com/sites/Finance
https://contoso.sharepoint.com/sites/HR
```

## Script parameters

Key parameters (defaults are defined in the script):

- `-CsvPath` path to CSV file (default `.\Sites.csv`)
- `-ConfigPath` path to JSON config (default `.\config.json`)
- `-SiteLogoPath` local image path for logo
- `-SiteThumbnailPath` local image path for site thumbnail
- `-HubSiteUrl` target hub URL (required by validation)

Other internal arrays/values in the script control:

- App IDs to install
- Page templates list
- Content type name and target lists

## How to run

From the `SiteProvisioning` folder:

```powershell
.\Set-SharePointSitePropertiesFromCsv.ps1 `
  -CsvPath ".\Sites.csv" `
  -ConfigPath ".\config.json" `
  -HubSiteUrl "https://contoso.sharepoint.com/sites/YourHub" `
  -SiteLogoPath "D:\Images\logo.png" `
  -SiteThumbnailPath "D:\Images\thumbnail.png"
```

## Notes

- The script requires `SiteUrl` in every row; blank rows are skipped.
- If `config.json` or CSV is missing/invalid, the script exits with an error.
- Some values are currently hardcoded in the script (for example, Entra group object ID, app IDs, content type/list names). Update those values to match your tenant standards before running in production.
