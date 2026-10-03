resource "azurerm_resource_group" "this" {
  name     = var.resource_group_name
  location = var.location
}

resource "azurerm_automation_account" "this" {
  name                = var.automation_account_name
  location            = azurerm_resource_group.this.location
  resource_group_name = azurerm_resource_group.this.name
  sku_name            = "Basic"

  identity {
    type = "SystemAssigned"
  }
}

resource "azurerm_automation_runbook" "get_groups" {
  name                    = "Get-Groups"
  location                = azurerm_resource_group.this.location
  resource_group_name     = azurerm_resource_group.this.name
  automation_account_name = azurerm_automation_account.this.name
  log_verbose             = true
  log_progress            = true
  description             = "Syncs Entra group names and selected group member emails into SharePoint choice fields."
  runbook_type = "PowerShell72"
  content      = replace(file("${path.module}/../AzureAutomation/Get-Groups.ps1"), "\uFEFF", "")

  # Provider reads PowerShell72 back as PowerShell. An in-place update then sends the
  # wrong type and Azure returns "Runbook Type cannot be modified."
  lifecycle {
    ignore_changes = [runbook_type, content]
  }
}

# Publish script content without changing runbook type.
resource "terraform_data" "get_groups_content" {
  input = replace(file("${path.module}/../AzureAutomation/Get-Groups.ps1"), "\uFEFF", "")

  depends_on = [azurerm_automation_runbook.get_groups]

  provisioner "local-exec" {
    interpreter = ["powershell", "-NoProfile", "-Command"]
    environment = {
      SCRIPT_CONTENT = self.input
    }
    command = <<-EOT
      $ErrorActionPreference = 'Stop'
      $tmp = Join-Path $env:TEMP 'Get-Groups-runbook.ps1'
      $utf8 = New-Object System.Text.UTF8Encoding $false
      [System.IO.File]::WriteAllText($tmp, $env:SCRIPT_CONTENT, $utf8)
      $token = az account get-access-token --resource https://management.azure.com --query accessToken --output tsv
      if ($LASTEXITCODE -ne 0) { throw "Could not get an Azure access token." }
      $base = "https://management.azure.com/subscriptions/${var.subscription_id}/resourceGroups/${var.resource_group_name}/providers/Microsoft.Automation/automationAccounts/${var.automation_account_name}/runbooks/Get-Groups"
      $headers = @{ Authorization = "Bearer $token" }
      Invoke-RestMethod -Method Put -Uri "$base/draft/content?api-version=2023-11-01" -Headers $headers -ContentType "text/powershell; charset=utf-8" -InFile $tmp
      Invoke-RestMethod -Method Post -Uri "$base/publish?api-version=2023-11-01" -Headers $headers
    EOT
  }
}

# Weekly at 07:00 New Zealand time, every day except Friday.
# start_time must be in the future on the first apply; later applies leave it alone.
resource "azurerm_automation_schedule" "get_groups" {
  name                    = "Get-Groups-0700-except-Friday"
  resource_group_name     = azurerm_resource_group.this.name
  automation_account_name = azurerm_automation_account.this.name
  frequency               = "Week"
  interval                = 1
  timezone                = "Pacific/Auckland"
  start_time              = var.schedule_start_time
  description             = "Run Get-Groups at 9:20am every day except Friday (New Zealand time)."
  week_days               = ["Monday", "Tuesday", "Wednesday", "Thursday", "Saturday"]

 
}

resource "azurerm_automation_job_schedule" "get_groups" {
  resource_group_name     = azurerm_resource_group.this.name
  automation_account_name = azurerm_automation_account.this.name
  schedule_name           = azurerm_automation_schedule.get_groups.name
  runbook_name            = azurerm_automation_runbook.get_groups.name

  parameters = {
    sharepoint_site_url = var.sharepoint_site_url
    managed_identity    = var.managed_identity_client_id
  }

  depends_on = [terraform_data.get_groups_content]

  # Replacing the runbook deletes this link in Azure. Recreate it when that happens.
  lifecycle {
    replace_triggered_by = [
      azurerm_automation_runbook.get_groups.id,
      terraform_data.get_groups_content
    ]
  }
}
