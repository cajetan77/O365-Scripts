output "automation_account_name" {
  value = azurerm_automation_account.this.name
}

output "resource_group_name" {
  value = azurerm_resource_group.this.name
}

output "principal_id" {
  description = "System-assigned managed identity object ID. Use this as $PrincipalId in Set-SystemManagedId.ps1."
  value       = azurerm_automation_account.this.identity[0].principal_id
}

output "runbook_name" {
  value = azurerm_automation_runbook.get_groups.name
}

output "schedule_name" {
  value = azurerm_automation_schedule.get_groups.name
}
