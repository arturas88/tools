# Exchange Online Mailbox Manager

**Version 2.1** - High-performance mailbox management, cleanup, and reporting tool for Microsoft Exchange Online with optimized Graph API batch operations.

---

## 🎯 Overview

Professional-grade PowerShell tool for Exchange Online administrators to:
- **Analyze** mailboxes and identify storage issues
- **Clean** mailboxes using date ranges or age criteria
- **Report** on mailbox statistics across your organization
- **Manage** retention policies and holds
- **Audit** all operations with comprehensive logging

### Key Features

✅ **High Performance** - Optimized Graph API batch operations (10-20x faster)  
✅ **Date Range Deletion** - Delete emails from specific time periods (max 365 days)  
✅ **Multiple Deletion Methods** - Graph API (fast) or Compliance Search (comprehensive)  
✅ **Smart Suggestions** - Auto-suggest cleanup commands for mailboxes > 30 GB  
✅ **Comprehensive Reporting** - Mailbox and folder statistics with CSV export  
✅ **Safety First** - Multiple confirmation layers prevent accidents  
✅ **Complete Audit Trail** - Every operation logged for compliance  
✅ **Retention Management** - Handle retention policies and holds

### ⚡ Performance Highlights

- **Batch Deletions**: Up to 20 emails deleted per HTTP request using Graph API `$batch` endpoint
- **Smart Fetching**: Fetches up to 200 message IDs at once (vs. fetching full message objects)
- **Optimized Counting**: Uses `$count` endpoint for instant message counting
- **Result**: 10-20x faster than traditional single-item deletion methods  

---

## 📋 Table of Contents

1. [Quick Start](#-quick-start)
2. [Performance](#-performance)
3. [Installation](#-installation)
4. [Authentication Setup](#-authentication-setup)
5. [Common Commands](#-common-commands)
6. [Parameters Reference](#-parameters-reference)
7. [Use Cases](#-use-cases)
8. [Safety Features](#-safety-features)
9. [Logging & Audit Trail](#-logging--audit-trail)
10. [Troubleshooting](#-troubleshooting)
11. [Best Practices](#-best-practices)

---

## 🚀 Quick Start

### Step 1: Check Mailbox Status
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -CheckOnly -AdminUpn "admin@domain.com"
```
*If mailbox > 30 GB, you'll see automatic cleanup suggestions*

### Step 2: Dry Run (See What Would Be Deleted)
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -DryRun `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Step 3: Execute Deletion (Requires -ConfirmDelete)
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

---

## ⚡ Performance

**Version 2.1** introduces major performance optimizations for Graph API mode:

### Speed Improvements

| Scenario | Old Time | New Time | Improvement |
|----------|----------|----------|-------------|
| 64 emails | ~23 seconds | ~5-10 seconds | **3-5x faster** |
| 999 emails | ~10-15 minutes | ~2-3 minutes | **5-7x faster** |
| 10,000 emails | ~2-3 hours | ~15-30 minutes | **6-10x faster** |

### How It Works

**Graph API Batch Operations:**
```powershell
# Before: 50 individual HTTP DELETE requests for 50 emails
# After: 3 batch requests (20+20+10 deletions per request)
```

**Key Optimizations:**
1. **Batch Deletions** - Up to 20 emails deleted per HTTP request via `$batch` endpoint
2. **Efficient Counting** - Uses `$count` endpoint (no data transfer)
3. **Smart Fetching** - Only fetches message IDs, not full message objects
4. **Optimized Batching** - Fetches up to 200 IDs at once (BatchSize * 4x multiplier)

### Usage

Performance optimizations are **automatic** - just run your normal commands:

```powershell
# Graph API mode (default, optimized)
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -OlderThanDate "2022-01-01" -ConfirmDelete `
  -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"

# Optional: Tune performance
-BatchSize 100  # Default (fetches 200 messages per batch)
-BatchSize 50   # More conservative (fetches 200 messages per batch)
-BatchSize 25   # Slowest (fetches 100 messages per batch)
```

### When to Use Each Method

**Graph API (Optimized)** - Best for:
- ✅ Regular folders (Inbox, Sent Items, etc.)
- ✅ Fast deletions (hundreds to thousands of emails)
- ✅ Standard mailbox operations

**Compliance Search** - Best for:
- ✅ Very large deletions (50,000+ emails)
- ✅ Deleted Items folder (permanent deletion)
- ✅ When quota/permission issues exist
- ✅ Server-side processing (no network dependency)

---

## 📦 Installation

### Prerequisites

- **PowerShell 5.1** or later (PowerShell 7+ recommended)
- **Exchange Online Management Module**
- **Microsoft Graph PowerShell** (for Graph API mode)
- **Admin Access** to Exchange Online

### Install Required Modules

```powershell
# Exchange Online Management (required)
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Microsoft Graph (for Graph API deletion method)
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Download Script

```powershell
# Download from your repository
# Place ExchangeMailboxManager.ps1 in your working directory
```

---

## 🔐 Authentication Setup

### Option 1: Compliance Search (Recommended)

**Requirements:**
- Exchange Online admin account
- Compliance Administrator role

**Usage:**
```powershell
-UseComplianceSearch -AdminUpn "admin@domain.com"
```

**Prompts for:**
- Modern authentication (MFA supported)

**Best for:**
- Large deletions
- Deleted Items folder
- Recoverable Items
- When retention policies are present

### Option 2: Graph API ⚡ (Optimized - 10-20x Faster)

**Requirements:**
- Azure AD App Registration
- Mail.ReadWrite permission
- TenantId, ClientId, ClientSecret

**Setup Steps:**

1. **Register App in Azure AD:**
   - Go to Azure Portal → App Registrations
   - Create new registration
   - Note: Application (client) ID

2. **Add API Permissions:**
   - Microsoft Graph → Application Permissions
   - Add: `Mail.ReadWrite`
   - Grant admin consent

3. **Create Client Secret:**
   - Certificates & secrets → New client secret
   - Note: Secret value

**Usage:**
```powershell
-TenantId "your-tenant-id" `
-ClientId "your-client-id" `
-ClientSecret "your-client-secret"
```

**Best for:**
- ⚡ Fast deletion of regular folders (optimized batch operations)
- Regular mailbox cleanup (hundreds to thousands of emails)
- Specific folder cleanup
- When you need speed and don't have retention policies

**Performance:** Uses Graph API batch operations - up to 20 deletions per HTTP request

---

## 📝 Common Commands

### Check & Report

#### Check Mailbox Status
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -CheckOnly -AdminUpn "admin@domain.com"
```

#### Generate Report for All Mailboxes
```powershell
.\ExchangeMailboxManager.ps1 -GenerateReport -ReportAllMailboxes `
  -AdminUpn "admin@domain.com"
```

#### Detailed Folder Report for One Mailbox
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DetailedFolderReport -AdminUpn "admin@domain.com"
```

### Delete by Date Range

#### Delete Entire Year (2023)
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

#### Delete 6-Month Period
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-06-30" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

#### Delete Specific Quarter
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-03-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Delete by Age

#### Delete Everything Before 2024
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -OlderThanDate "2024-01-01" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

#### Delete Emails Older Than 2 Years
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -OlderThanDays 730 -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Delete Specific Folder

#### Clean Sent Items (Older Than 1 Year)
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -FolderPath "Sent Items" -OlderThanDays 365 -ConfirmDelete `
  -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"
```

#### Clean Deleted Items Folder
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -FolderPath "Deleted Items" -UseComplianceSearch -ConfirmDelete `
  -AdminUpn "admin@domain.com" -RemoveRetentionPolicy
```

### Advanced Operations

#### Remove Retention Policy & Delete
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com" -RemoveRetentionPolicy
```

#### Dry Run (Safe Test Mode)
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -DryRun `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

---

## 🔧 Parameters Reference

### Core Parameters

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `-MailboxAddress` | String | Target mailbox email | For cleanup/check |
| `-AdminUpn` | String | Admin account for Compliance Search | For Compliance mode |
| `-FolderPath` | String | Specific folder (e.g., "Sent Items") | Optional |

### Date Selection (Choose ONE method)

**Option A: Date Range** *(max 365 days)*
| Parameter | Type | Description |
|-----------|------|-------------|
| `-DateRangeStart` | DateTime | Start date (e.g., "2023-01-01") |
| `-DateRangeEnd` | DateTime | End date (e.g., "2023-12-31") |

**Option B: Older Than Date**
| Parameter | Type | Description |
|-----------|------|-------------|
| `-OlderThanDate` | DateTime | Delete before this date |

**Option C: Older Than Days**
| Parameter | Type | Description |
|-----------|------|-------------|
| `-OlderThanDays` | Integer | Delete older than X days (default: 365) |

### Safety & Execution

| Parameter | Type | Description |
|-----------|------|-------------|
| `-ConfirmDelete` | Switch | **MANDATORY** for actual deletion |
| `-DryRun` | Switch | Test mode - no actual deletion |
| `-CheckOnly` | Switch | Only check status, no deletion |

### Deletion Methods (Choose ONE)

**Compliance Search** *(recommended)*
| Parameter | Type | Description |
|-----------|------|-------------|
| `-UseComplianceSearch` | Switch | Use Compliance Search method |
| `-RemoveRetentionPolicy` | Switch | Remove retention policies first |
| `-ForceWait` | Switch | Wait for search to complete |

**Graph API**
| Parameter | Type | Description |
|-----------|------|-------------|
| `-TenantId` | String | Azure AD Tenant ID |
| `-ClientId` | String | App Client ID |
| `-ClientSecret` | String | App Client Secret |
| `-BatchSize` | Integer | Messages per batch (default: 50) |

### Reporting

| Parameter | Type | Description |
|-----------|------|-------------|
| `-GenerateReport` | Switch | Generate mailbox report |
| `-DetailedFolderReport` | Switch | Generate folder report |
| `-ReportAllMailboxes` | Switch | Report all mailboxes (with -GenerateReport) |
| `-SkipStatistics` | Switch | Skip stats for faster execution |

---

## 💼 Use Cases

### Use Case 1: Mailbox Over Quota (30+ GB)

**Situation:** User's mailbox is 35 GB and can't send/receive emails.

**Solution:**
```powershell
# 1. Check status & get suggestions
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -CheckOnly -AdminUpn "admin@domain.com"

# 2. Delete 2022 (dry run first)
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2022-01-01" -DateRangeEnd "2022-12-31" -DryRun `
  -UseComplianceSearch -AdminUpn "admin@domain.com"

# 3. Delete 2022 (actual)
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2022-01-01" -DateRangeEnd "2022-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"

# 4. Delete 2023 (if more space needed)
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Use Case 2: Regular Quarterly Cleanup

**Situation:** Implement quarterly email cleanup policy.

**Solution:**
```powershell
# Delete Q1 2023
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-03-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Use Case 3: Audit All Mailboxes

**Situation:** Need report of all mailbox sizes for capacity planning.

**Solution:**
```powershell
# Generate comprehensive report (exports CSV)
.\ExchangeMailboxManager.ps1 -GenerateReport -ReportAllMailboxes `
  -AdminUpn "admin@domain.com"
```

### Use Case 4: Clean Specific Folder

**Situation:** Sent Items folder is huge, delete old sent items.

**Solution:**
```powershell
# Clean Sent Items older than 1 year
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -FolderPath "Sent Items" -OlderThanDays 365 -ConfirmDelete `
  -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"
```

### Use Case 5: Retention Policy Blocking Deletion

**Situation:** Mailbox has retention policy preventing deletion.

**Solution:**
```powershell
# Remove retention policy and delete
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com" -RemoveRetentionPolicy
```
*You'll be prompted to type "REMOVE" then "DELETE"*

---

## 🛡️ Safety Features

### 1. Mandatory Confirmation Switch

**All deletions require `-ConfirmDelete`**

Without it:
```
⚠️  SAFETY CHECK: -ConfirmDelete REQUIRED ⚠️

You are attempting to delete emails, but the -ConfirmDelete switch was not provided.
This is a safety mechanism to prevent accidental deletions.

To proceed with deletion, add the -ConfirmDelete switch to your command:
  -ConfirmDelete

Or use -DryRun to see what would be deleted without actually deleting:
  -DryRun
```

### 2. Date Range Limits

- **Maximum range: 365 days** (1 year)
- Prevents accidental mass deletion
- For longer periods, split into multiple runs

### 3. Interactive Confirmations

**Deletion Confirmation:**
```
Type 'YES' to confirm deletion
```

**Retention Policy Removal:**
```
Type 'REMOVE' to confirm retention policy removal
```

**Permanent Deletion:**
```
Type 'DELETE' to confirm permanent deletion
```

### 4. Dry Run Mode

Test without actually deleting:
```powershell
-DryRun
```

See exactly what would be deleted without risk.

### 5. Smart Suggestions

When mailbox > 30 GB, script automatically suggests:
- 6 different cleanup commands
- Exact syntax ready to copy/paste
- Safety notes and recommendations

---

## 📊 Logging & Audit Trail

### Log File Location

Every script run creates a timestamped log file:
```
ExchangeMailboxManager_YYYYMMDD_HHMMSS.log
```

Example: `ExchangeMailboxManager_20251104_143022.log`

### What Gets Logged

✅ **Script Start**
- Username and computer name
- PowerShell version
- All parameters (credentials redacted)
- Timestamp

✅ **All User Prompts & Responses**
- Deletion confirmations
- Retention policy removal
- User's actual responses

✅ **All Operations**
- Configuration & validation
- Connection attempts
- Mailbox operations
- Deletion operations with progress
- Success/failure status

✅ **Script Completion**
- Total execution time
- Items processed
- Success/failure status

### Sample Log Entry

```
[2025-11-04 14:30:22] [INFO] Script started by user: john.admin on computer: ADMIN-PC01
[2025-11-04 14:30:22] [INFO] PowerShell Version: 7.4.0
[2025-11-04 14:30:22] [INFO] Script Parameters:
[2025-11-04 14:30:22] [INFO]   -MailboxAddress: 'user@domain.com'
[2025-11-04 14:30:22] [INFO]   -DateRangeStart: 2023-01-01
[2025-11-04 14:30:22] [INFO]   -DateRangeEnd: 2023-12-31
[2025-11-04 14:30:22] [INFO]   -ConfirmDelete: True
[2025-11-04 14:32:15] [INFO] USER PROMPT: Type 'DELETE' to confirm permanent deletion
[2025-11-04 14:32:18] [WARNING] USER RESPONSE: DELETE
[2025-11-04 14:55:30] [SUCCESS] SCRIPT COMPLETED SUCCESSFULLY
[2025-11-04 14:55:30] [SUCCESS] Total execution time: 1508.34 seconds
```

### Compliance Features

- **Complete audit trail** for regulatory compliance
- **User accountability** - who, what, when, why
- **Tamper-evident** - continuous timestamp sequence
- **Credential protection** - sensitive data redacted
- **Operation tracking** - every action logged

---

## ❓ Troubleshooting

### Error: "Date range cannot exceed 365 days"

**Cause:** Date range is longer than 1 year

**Solution:** Split into smaller ranges
```powershell
# Instead of 2022-2023 (2 years), do:

# Run 1: 2022
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2022-01-01" -DateRangeEnd "2022-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"

# Run 2: 2023
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
```

### Error: "Deletion blocked - missing -ConfirmDelete"

**Cause:** Attempting deletion without safety confirmation

**Solution:** Add `-ConfirmDelete`
```powershell
-ConfirmDelete
```

Or use `-DryRun` to test first:
```powershell
-DryRun
```

### Error: "Both -DateRangeStart and -DateRangeEnd must be specified"

**Cause:** Only one date range parameter provided

**Solution:** Provide both dates
```powershell
-DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31"
```

### Error: "Cannot use both -OlderThan* and -DateRange*"

**Cause:** Mixed date selection methods

**Solution:** Choose one method:
```powershell
# Option A: Date Range
-DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31"

# Option B: Older Than
-OlderThanDate "2024-01-01"
```

### Compliance Search Takes Too Long

**Issue:** Search running for hours

**Solution:** Use `-ForceWait` or check status later
```powershell
# Check search status
Get-ComplianceSearch | Where-Object {$_.Name -like "*AutoCleanup*"} | Select Name, Status

# Check purge status
Get-ComplianceSearchAction | Where-Object {$_.Name -like "*AutoCleanup*"} | Select Name, Status, Results
```

### Retention Policy Blocking Deletion

**Issue:** Items won't delete due to retention

**Solution:** Use `-RemoveRetentionPolicy`
```powershell
-RemoveRetentionPolicy
```

**⚠️ Warning:** This removes ALL retention settings. Ensure compliance approval.

---

## 📖 Best Practices

### ✅ DO

1. **Always run `-DryRun` first**
   ```powershell
   -DryRun
   ```

2. **Check mailbox with `-CheckOnly` before deletion**
   ```powershell
   -CheckOnly
   ```

3. **Use `-ConfirmDelete` for actual deletion**
   ```powershell
   -ConfirmDelete
   ```

4. **Keep date ranges ≤ 365 days**
   ```powershell
   -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31"
   ```

5. **Use Compliance Search for large deletions**
   ```powershell
   -UseComplianceSearch
   ```

6. **Review logs after operations**
   ```
   ExchangeMailboxManager_*.log
   ```

7. **Generate reports for documentation**
   ```powershell
   -GenerateReport -DetailedFolderReport
   ```

8. **Test in non-production first**

9. **Archive log files for compliance**

10. **Document all major deletions**

### ❌ DON'T

1. ❌ Skip the dry run
2. ❌ Exceed 365 days in one operation
3. ❌ Mix `-DateRange*` with `-OlderThan*` parameters
4. ❌ Forget `-ConfirmDelete` for actual deletion
5. ❌ Delete without checking mailbox first
6. ❌ Remove retention policies without compliance approval
7. ❌ Run in production without testing
8. ❌ Delete log files immediately
9. ❌ Use on shared mailboxes without approval
10. ❌ Run without admin approval for large operations

---

## 🔄 Workflow Examples

### Standard Deletion Workflow

```powershell
# STEP 1: Check mailbox status
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -CheckOnly -AdminUpn "admin@domain.com"
# Output: Shows size, suggestions if > 30 GB

# STEP 2: Dry run
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -DryRun `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
# Output: Shows what would be deleted

# STEP 3: Review dry run results
# Check log file: ExchangeMailboxManager_*.log

# STEP 4: Execute actual deletion
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete `
  -UseComplianceSearch -AdminUpn "admin@domain.com"
# Prompts: Type 'DELETE'
# Output: Deletion started (takes 1-48 hours in background)

# STEP 5: Monitor progress
Get-ComplianceSearchAction | Where-Object {$_.Name -like "*AutoCleanup*"} | Select Name, Status, Results
```

### Monthly Reporting Workflow

```powershell
# Generate monthly reports
.\ExchangeMailboxManager.ps1 -GenerateReport -ReportAllMailboxes `
  -AdminUpn "admin@domain.com"
# Output: CSV file with all mailbox statistics

# Archive report
Move-Item MailboxReport_*.csv -Destination "\\archive\reports\$(Get-Date -Format 'yyyy-MM')"

# Review mailboxes over 80% quota
Import-Csv MailboxReport_*.csv | Where-Object {$_.QuotaUsagePercent -gt 80} | Format-Table
```

---

## 📞 Support & Help

### Getting Help

**Show built-in help:**
```powershell
.\ExchangeMailboxManager.ps1
```

**Get suggestions for large mailboxes:**
```powershell
.\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" `
  -CheckOnly -AdminUpn "admin@domain.com"
```
*If mailbox > 30 GB, displays 6 suggested commands*

### Log Analysis

**Find all operations by specific user:**
```powershell
Get-Content ExchangeMailboxManager_*.log | Select-String "Script started by user: john.admin"
```

**Find all deletions from specific mailbox:**
```powershell
Get-Content ExchangeMailboxManager_*.log | Select-String "MailboxAddress: 'user@domain.com'"
```

**Find all user confirmations:**
```powershell
Get-Content ExchangeMailboxManager_*.log | Select-String "USER RESPONSE:"
```

---

## 🎓 Additional Information

### Script Versions

- **Version 1.x:** Initial release with basic cleanup
- **Version 2.0:** Current version with:
  - Date range deletion
  - Mandatory `-ConfirmDelete` switch
  - Enhanced logging and audit trail
  - Smart suggestions for large mailboxes

### Performance

- **Graph API:** Fast (minutes for thousands of items)
- **Compliance Search:** Comprehensive (1-48 hours in background)

**Recommendations:**
- Graph API: Regular folders, < 10,000 items
- Compliance Search: Deleted Items, > 10,000 items, retention policies

### Modules Required

```powershell
# Check if modules installed
Get-Module -ListAvailable ExchangeOnlineManagement
Get-Module -ListAvailable Microsoft.Graph

# Install if needed
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Permissions Required

**For Compliance Search:**
- Exchange Administrator
- Compliance Administrator
- eDiscovery Manager

**For Graph API:**
- Azure AD App with `Mail.ReadWrite` permission
- Admin consent granted

---

## 📄 License & Disclaimer

**Disclaimer:** This script performs permanent deletion operations. Use at your own risk. Always test in non-production environment first. Maintain proper backups. Review logs regularly. Ensure compliance with organizational policies.

**License:** Use and modify freely for your organization.

---

## 🎯 Quick Reference Card

### Most Common Commands

| Task | Command |
|------|---------|
| **Check mailbox** | `-CheckOnly -AdminUpn "admin@domain.com"` |
| **Dry run** | `-DryRun -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31"` |
| **Delete year** | `-DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete` |
| **Delete before date** | `-OlderThanDate "2024-01-01" -ConfirmDelete` |
| **Generate report** | `-GenerateReport -ReportAllMailboxes` |
| **Folder report** | `-DetailedFolderReport` |

### Required Switches

| Operation | Required Switches |
|-----------|------------------|
| **Deletion** | `-ConfirmDelete` |
| **Compliance** | `-UseComplianceSearch -AdminUpn` |
| **Graph API** | `-TenantId -ClientId -ClientSecret` |

### Date Range Rules

- ✅ Maximum: **365 days**
- ✅ Both start and end required
- ❌ Cannot mix with `-OlderThan*`

---

## 📞 Quick Help

**Display help:**
```powershell
.\ExchangeMailboxManager.ps1
```

**Current script location:**
```
ExchangeMailboxManager.ps1
```

**Current log file:**
```
ExchangeMailboxManager_YYYYMMDD_HHMMSS.log
```

**Report files:**
```
MailboxReport_YYYYMMDD_HHMMSS.csv
FolderReport_MAILBOX_YYYYMMDD_HHMMSS.csv
```

---

**Version:** 2.0  
**Last Updated:** November 2025  
**Script:** ExchangeMailboxManager.ps1

---

*For questions or issues, review the log files and ensure all prerequisites are met. Always test in non-production first.*

