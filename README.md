# Microsoft 365 PowerShell Scripts

A small collection of PowerShell scripts for administering Microsoft 365 (SharePoint Online / Microsoft 365 Groups / Teams), built around **PnP.PowerShell** and Microsoft Graph.

> ⚠️ **Warning:** Always test in a non-production tenant first.

---

## What’s in this repo

Typical tasks these scripts cover:
- SharePoint Online site admin tasks (including tenant admin actions)
- Microsoft 365 Group–connected site operations
- Recycle bin / deleted site cleanup

---

## Prerequisites

### 1) PowerShell
- PowerShell 7+

Check:
```powershell
$PSVersionTable.PSVersion

### 2) Install Pnp Powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Verify:
Get-Module PnP.PowerShell -ListAvailable | Select-Object Name, Version

### 3) Install Microsoft Graph
Install-Module Microsoft.Graph -Scope CurrentUser
Verify:
Get-Module Microsoft.Graph -ListAvailable | Select Name,Version
