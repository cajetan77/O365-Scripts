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
