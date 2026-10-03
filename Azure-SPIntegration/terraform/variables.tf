variable "subscription_id" {
  type        = string
  description = "Azure subscription ID to deploy into."
}

variable "resource_group_name" {
  type        = string
  description = "Resource group for the Automation Account."
  default     = "SPO-Automation"
}

variable "location" {
  type        = string
  description = "Azure region."
  default     = "australiasoutheast"
}

variable "automation_account_name" {
  type        = string
  description = "Automation Account name. Must be unique in the resource group."
  default     = "aa-spo-automation12"
}

variable "sharepoint_site_url" {
  type        = string
  description = "Passed to the runbook parameter SharePointSiteUrl."
  default     = "https://caje77sharepoint.sharepoint.com/sites/CajIntra/"
}

variable "schedule_start_time" {
  type        = string
  description = "First run of the schedule, RFC3339, at 07:00 New Zealand time on a non-Friday. Must be at least 5 minutes in the future on the first apply."
  default     = "2026-10-04T13:00:00+13:00"
}

variable "managed_identity_client_id" {
  type        = string
  description = "Passed to the runbook parameter ManagedIdentity (user-assigned managed identity client ID)."
  default     = "66a1852a-1f21-46a2-ad58-35fc4c3f1530"
}
