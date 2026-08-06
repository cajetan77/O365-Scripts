<#
.SYNOPSIS
    Syncs Entra DLs / mail-enabled security groups into SharePoint choice fields.

.DESCRIPTION
    Ensures the SharePoint list "Groups" exists (creates it if missing), ensures the
    choice fields exist, then syncs Entra mail-enabled groups into those fields:
      - "Distribution List"  <- mail-enabled, non-security groups (classic DLs)
      - "Mail Enabled"       <- mail-enabled security groups (not Microsoft 365 groups)

.NOTES
    Azure Automation setup
    ----------------------
    1. System-assigned Managed Identity on the Automation Account.
    2. Application permissions (admin consent):
         Group.Read.All
         Sites.ReadWrite.All (or SharePoint Sites.FullControl.All)
    3. Runtime modules (same Graph version, e.g. 2.38.1):
         Microsoft.Graph.Authentication
         Microsoft.Graph.Groups
         PnP.PowerShell
       Import PnP only after Graph work — loading both together can break Managed Identity (MSAL conflict).
    4. Automation variable: SHAREPOINT_SITE_URL
#>
[CmdletBinding()]
Param(
    [string]$ListTitle = 'Groups',
    [string]$DistributionListField = 'Distribution List',
    [string]$DistributionListInternalName = 'DistributionList',
    [string]$MailEnabledField = 'Mail Enabled',
    [string]$MailEnabledInternalName = 'MailEnabled'
)

$ErrorActionPreference = 'Stop'

$SharePointSiteUrl = Get-AutomationVariable -Name 'SHAREPOINT_SITE_URL' -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
    throw 'Automation variable SHAREPOINT_SITE_URL is required (e.g. https://contoso.sharepoint.com/sites/IT).'
}

function Write-RunbookLog {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    # Use Write-Host so log lines do not become function return values (Write-Output would).
    Write-Host "[$stamp] $Message"
}

function Import-MatchingGraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Groups'
    )

    $authVersions = @(
        Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' |
            Select-Object -ExpandProperty Version -Unique |
            Sort-Object -Descending
    )

    if ($authVersions.Count -eq 0) {
        throw 'Microsoft.Graph.Authentication is not installed in this Automation Account.'
    }

    $commonVersion = $null
    foreach ($version in $authVersions) {
        $missing = @()
        foreach ($moduleName in $requiredModules) {
            $match = Get-Module -ListAvailable -Name $moduleName |
                Where-Object { $_.Version -eq $version } |
                Select-Object -First 1
            if (-not $match) {
                $missing += $moduleName
            }
        }

        if ($missing.Count -eq 0) {
            $commonVersion = $version
            break
        }
    }

    if (-not $commonVersion) {
        $installed = foreach ($moduleName in $requiredModules) {
            $versions = @(
                Get-Module -ListAvailable -Name $moduleName |
                    Select-Object -ExpandProperty Version -Unique |
                    Sort-Object -Descending
            )
            "  $moduleName : $(if ($versions) { $versions -join ', ' } else { '(not installed)' })"
        }

        throw @"
Microsoft Graph modules must all be the SAME version in Azure Automation.
Install matching versions of:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.Groups

Currently available:
$($installed -join [Environment]::NewLine)
"@
    }

    foreach ($moduleName in $requiredModules) {
        Import-Module $moduleName -RequiredVersion $commonVersion -Force -ErrorAction Stop
    }

    Write-RunbookLog "Loaded Microsoft Graph modules version $commonVersion"
}

function Test-IsUnifiedGroup {
    param($Group)
    return (@($Group.GroupTypes) -contains 'Unified')
}

function Initialize-GroupsList {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$DistributionListField,
        [Parameter(Mandatory)][string]$DistributionListInternalName,
        [Parameter(Mandatory)][string]$MailEnabledField,
        [Parameter(Mandatory)][string]$MailEnabledInternalName
    )

    $list = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-RunbookLog "Creating list '$ListTitle'..."
        $null = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch -ErrorAction Stop
        Write-RunbookLog "Created list '$ListTitle'."
    }
    else {
        Write-RunbookLog "List '$ListTitle' already exists."
    }

    Initialize-ChoiceField -ListTitle $ListTitle -DisplayName $DistributionListField -InternalName $DistributionListInternalName
    Initialize-ChoiceField -ListTitle $ListTitle -DisplayName $MailEnabledField -InternalName $MailEnabledInternalName
}

function Resolve-PnPListField {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$InternalName,
        [Parameter(Mandatory)][string]$DisplayName
    )

    $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        $field = Get-PnPField -List $ListTitle -Identity $DisplayName -ErrorAction SilentlyContinue
    }
    if (-not $field) {
        throw "Choice field not found on list '$ListTitle'. Tried '$InternalName' and '$DisplayName'."
    }
    return $field
}

function Initialize-ChoiceField {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$InternalName
    )

    $field = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        $field = Get-PnPField -List $ListTitle -Identity $DisplayName -ErrorAction SilentlyContinue
    }

    if (-not $field) {
        Write-RunbookLog "Creating choice field '$DisplayName' ($InternalName) on list '$ListTitle'..."
        Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Choice -AddToDefaultView -ErrorAction Stop | Out-Null
        Write-RunbookLog "Created field '$DisplayName'."
        return
    }

    Write-RunbookLog "Field '$DisplayName' already exists (InternalName=$($field.InternalName))."
}

function Set-PnPChoiceFieldValues {
    param(
        [Parameter(Mandatory)][string]$ListTitle,
        [Parameter(Mandatory)][string]$InternalName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Choices
    )

    $uniqueChoices = @(
        $Choices |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Sort-Object -Unique
    )

    $field = Resolve-PnPListField -ListTitle $ListTitle -InternalName $InternalName -DisplayName $DisplayName
    [xml]$schemaXml = $field.SchemaXml

    $choicesNode = $schemaXml.SelectSingleNode('/Field/CHOICES')
    if (-not $choicesNode) {
        $choicesNode = $schemaXml.CreateElement('CHOICES')
        [void]$schemaXml.DocumentElement.AppendChild($choicesNode)
    }
    else {
        $choicesNode.RemoveAll()
    }

    foreach ($choice in $uniqueChoices) {
        $choiceNode = $schemaXml.CreateElement('CHOICE')
        $choiceNode.InnerText = $choice
        [void]$choicesNode.AppendChild($choiceNode)
    }

    if ($schemaXml.Field.HasAttribute('FillInChoice')) {
        $schemaXml.Field.SetAttribute('FillInChoice', 'FALSE')
    }
    if ($schemaXml.Field.HasAttribute('Format')) {
        $schemaXml.Field.SetAttribute('Format', 'Dropdown')
    }

    Set-PnPField -List $ListTitle -Identity $field.InternalName -Values @{ SchemaXml = $schemaXml.OuterXml } -ErrorAction Stop
    Write-RunbookLog "Updated '$DisplayName' ($($field.InternalName)) with $($uniqueChoices.Count) choice(s)."
}

try {
    # Import Graph modules first. Do NOT import PnP.PowerShell yet — it ships a
    # different Microsoft.Identity.Client and breaks Connect-MgGraph -Identity.
    Import-MatchingGraphModules

    Write-RunbookLog 'Connecting to Microsoft Graph with Managed Identity...'
    Connect-MgGraph -Identity -NoWelcome

    Write-RunbookLog 'Retrieving groups...'
    $allGroups = @(Get-MgGroup -All -Property Id, DisplayName, MailEnabled, SecurityEnabled, GroupTypes, Mail)

    $distributionGroups = @(
        $allGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            $_.SecurityEnabled -eq $false -and
            -not (Test-IsUnifiedGroup -Group $_)
        }
    )

    $mailEnabledSecurityGroups = @(
        $allGroups | Where-Object {
            $_.MailEnabled -eq $true -and
            $_.SecurityEnabled -eq $true -and
            -not (Test-IsUnifiedGroup -Group $_)
        }
    )

    Write-RunbookLog "Found $($distributionGroups.Count) classic distribution list(s) and $($mailEnabledSecurityGroups.Count) mail-enabled security group(s)."

    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    Import-Module PnP.PowerShell -ErrorAction Stop

    Write-RunbookLog "Connecting to SharePoint: $SharePointSiteUrl"
    Connect-PnPOnline -Url $SharePointSiteUrl -ManagedIdentity

    Initialize-GroupsList `
        -ListTitle $ListTitle `
        -DistributionListField $DistributionListField `
        -DistributionListInternalName $DistributionListInternalName `
        -MailEnabledField $MailEnabledField `
        -MailEnabledInternalName $MailEnabledInternalName

    Set-PnPChoiceFieldValues `
        -ListTitle $ListTitle `
        -InternalName $DistributionListInternalName `
        -DisplayName $DistributionListField `
        -Choices @($distributionGroups.DisplayName)

    Set-PnPChoiceFieldValues `
        -ListTitle $ListTitle `
        -InternalName $MailEnabledInternalName `
        -DisplayName $MailEnabledField `
        -Choices @($mailEnabledSecurityGroups.DisplayName)

    Write-RunbookLog 'Group choice sync complete.'
}
catch {
    Write-RunbookLog "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
