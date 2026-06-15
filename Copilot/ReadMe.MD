# Microsoft 365 Copilot Audit and Investigation Guide

## Table of Contents

1. Overview
2. Copilot Interaction Architecture
3. Audit Log Investigation
4. Exporting Copilot Interaction Logs
5. PowerShell Export Examples
6. Understanding Audit Records
7. Reporting Examples
8. What Audit Logs Can and Cannot Tell You
9. eDiscovery Setup
10. Copilot eDiscovery Searches
11. Review Sets
12. Investigation Scenarios
13. Time Zone Considerations
14. Governance Recommendations
15. References

---

# 1. Overview

This document describes how to:

* Export Microsoft 365 Copilot Interaction logs
* Analyse Copilot adoption
* Investigate Copilot activity
* Search Copilot conversations using eDiscovery
* Build governance and compliance reporting

---

# 2. Copilot Interaction Architecture

## High-Level Flow

User
→ Microsoft 365 Copilot
→ Grounding via Microsoft Graph
→ Referenced Content
→ LLM Processing
→ Response Generated
→ Audit Record Created


---

# 3. Audit Log Investigation

Audit logs answer:

* Who used Copilot?
* When was Copilot used?
* Which application was used?
* Which resources were referenced?

Audit logs DO NOT contain:

* User prompts
* Copilot responses
* Full conversations

---

# 4. Exporting Copilot Interaction Logs

## Purview Portal

Purview
→ Audit
→ Search

Activity:

CopilotInteraction

Export:

Download all results

---


# 6. Understanding Audit Records

Common Fields:

| Field             | Purpose            |
| ----------------- | ------------------ |
| UserId            | User               |
| Operation         | Activity           |
| Workload          | Copilot            |
| AppHost           | Application        |
| AppIdentity       | Copilot Service    |
| AccessedResources | Referenced Content |
| CreationDate      | Timestamp          |

---

# 7. Reporting Examples

## Top Users

* Interaction counts
* Adoption metrics


---

# 8. What Audit Logs Can and Cannot Tell You

## Audit Can Tell You

✓ User

✓ Time

✓ Application

✓ Referenced Files

✓ Interaction Count

## Audit Cannot Tell You

✗ Prompt Text

✗ Response Text

✗ Generated Code

✗ Conversation Context

Use eDiscovery for these.

---

# 9. eDiscovery Setup

## Prerequisites

### Roles

* eDiscovery Manager
* eDiscovery Administrator
* Compliance Administrator


#### eDiscovery Premium

* Review Sets



---

## Create a Case

Purview
→ eDiscovery
→ Cases
→ Create Case

Example:

Copilot Investigation - June 2026

---

## Create Search

Select:

Exchange Mailboxes

Copilot conversations are stored within Exchange Online mailboxes.

---

# 10. Copilot eDiscovery Searches

## All Copilot Conversations

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.*
```

## Copilot Chat

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.BizChat
```

## Word Copilot

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.Word
```

## Excel Copilot

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.Excel
```

## PowerPoint Copilot

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.PowerPoint
```

## Outlook Copilot

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.Outlook
```

## Teams Copilot

```text
ItemClass:IPM.SkypeTeams.Message.Copilot.Teams
```

---

# 11. Review Sets

Review Sets are available in eDiscovery Premium.

Purpose:

Search
→ Collect
→ Filter
→ Tag
→ Export

## Benefits

* Analytics
* Deduplication
* Tagging
* Faster searching

---

## Suggested Tags

| Tag             | Purpose                  |
| --------------- | ------------------------ |
| Code Generation | Code Requests            |
| Security        | Security Related         |
| Sensitive Data  | Confidential Information |
| HR              | Employee Information     |
| Privacy         | Personal Data            |
| False Positive  | Excluded Results         |

---

# 12. Investigation Scenarios

## Code Generation

Search:

```text
PowerShell
Python
JavaScript
SQL
C#
code snippet
generate code
```

---

## Sensitive Information

Search:

```text
salary
payroll
passport
confidential
restricted
bank account
```

---

## SharePoint Administration

Search:

```text
SharePoint
PnP
SPFx
Power Automate
Azure
```

---

## Copilot Adoption

Use Audit Logs

Questions:

* Who uses Copilot?
* Which apps are used?
* How often?

---

## Content Investigation

Use eDiscovery

Questions:

* What was asked?
* What was generated?
* Was sensitive information involved?

---

# 13. Time Zone Considerations

Audit and eDiscovery exports commonly use UTC.

Example:

```text
2026-06-10T01:30:00Z
```

New Zealand:

NZST = UTC +12

NZDT = UTC +13

Convert before reporting.

---

# 14. Governance Recommendations

## Stage 1

Audit Reporting

Track:

* Users
* Departments
* Applications
* Interaction Counts

---

## Stage 2

Risk Monitoring

Implement:

* Sensitivity Labels
* DLP Policies
* Insider Risk Management

---

## Stage 3

Investigations

Use eDiscovery

Use Review Sets

Export only relevant findings.

---

# 15. References

Microsoft Purview Audit Copilot Documentation

https://learn.microsoft.com/en-us/purview/audit-copilot

Investigating Microsoft 365 Copilot Interactions

https://nikkichapple.com/investigating-microsoft-365-copilot-interactions/

Microsoft Security Community Webinar

https://www.youtube.com/watch?v=LZZphUPvYr4
