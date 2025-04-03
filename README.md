# SharePoint Online & OneDrive Folder Activity Monitor

![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=for-the-badge&logo=powershell&logoColor=white)
![Microsoft SharePoint](https://img.shields.io/badge/Microsoft_SharePoint-0078D4?style=for-the-badge&logo=microsoft-sharepoint&logoColor=white)
![Microsoft OneDrive](https://img.shields.io/badge/Microsoft_OneDrive-0078D4?style=for-the-badge&logo=microsoft-onedrive&logoColor=white)

A PowerShell script to audit and report on folder-level activities in SharePoint Online and OneDrive for Business.

## Features

- 🔍 **Comprehensive Activity Tracking**:
  - Folder creations, modifications, and deletions
  - Recycle bin activities (1st & 2nd stage)
  - Restoration events
- 🔐 **Multiple Authentication Methods**:
  - Interactive login (MFA supported)
  - Certificate-based authentication
  - Service account credentials
- 📊 **Advanced Filtering**:
  - Site-specific auditing
  - User-specific activity tracking
  - Workload separation (SPO/OneDrive)
- 📁 **Enhanced CSV Export**:
  - Risk scoring for sensitive activities
  - Duration since activity
  - Automatic file opening option

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online PowerShell V2 module
- One of these roles:
  - Global Administrator
  - SharePoint Administrator + Compliance Administrator

## Installation

1. Clone the repository:
   ```powershell
   git clone https://github.com/RapidScripter/sharepoint-folder-activity-monitor.git
   cd sharepoint-folder-activity-monitor

2. Install the required module:
   ```powershell
   Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
   ```

## Usage

### Basic Commands

```powershell
# Interactive MFA session (last 180 days)
.\Monitor-FolderActivities.ps1

# Custom date range
.\Monitor-FolderActivities.ps1 -StartDate "2025-01-01" -EndDate "2025-01-31"

# SharePoint-only report
.\Monitor-FolderActivities.ps1 -SharePointOnline

# OneDrive-only audit
.\Monitor-FolderActivities.ps1 -OneDrive
```

### Advanced Options

| Parameter               | Description                          | Example                           |
|-------------------------|--------------------------------------|-----------------------------------|
| `-StartDate`            | Report start date                    | `-StartDate "2025-01-01"`         |
| `-EndDate`              | Report end date                      | `-EndDate "2025-01-31"`           |
| `-PerformedBy`          | Filter by specific user              | `-PerformedBy "user@domain.com"`  |
| `-SiteUrl`              | Audit single site                    | `-SiteUrl "https://..."`          |
| `-ImportSitesCsv`       | Bulk audit from CSV                  | `-ImportSitesCsv "sites.csv"`     |
| `-SharePointOnline`     | SharePoint-only report               | `-SharePointOnline`               |
| `-OneDrive`             | OneDrive-only report                 | `-OneDrive`                       |
| `-ClientId`             | App ID for certificate auth          | `-ClientId "xxxxxxxx-xxxx..."`    |
| `-CertificateThumbprint`| Certificate thumbprint for auth      | `-CertificateThumbprint "A1B2..."`|
| `-OutputPath`           | Custom report directory              | `-OutputPath "C:\AuditReports"`   |
| `-ThrottleLimit`        | Parallel processing threads (1-100)  | `-ThrottleLimit 10`               |

## Output

The script generates a CSV report with these columns:

- **Activity Time**: When activity occurred
- **Activity**: Type of folder operation
- **Folder Name**: Name of affected folder
- **Performed By**: User who performed action
- **Folder URL**: Full path to folder  
- **Site URL**: Parent site URL
- **Workload**: SharePoint or OneDrive
- **Risk Score**: High/Medium/Low classification
- **Duration (Days)**: Days since activity
- **More Info**: Full audit details (JSON)

Sample output filename: `Audit_SPO_Folder_Activity_Report_2023-08-15_143022.csv`

## Example Output

| Activity Time       | Activity          | Folder Name     | Performed By       | Risk Score | Site URL                     |
|---------------------|-------------------|-----------------|--------------------|------------|------------------------------|
| 2023-08-15 09:30:22 | FolderCreated     | Project Docs    | user1@domain.com   | Low        | https://.../sites/marketing  |
| 2023-08-15 10:15:41 | FolderDeleted     | Budget 2023     | external@vendor.com| High       | https://.../sites/finance    |
| 2023-08-15 11:20:33 | FolderRestored    | Client Files    | admin@domain.com   | Medium     | https://.../sites/legal      |

## Troubleshooting

| Error/Symptom | Solution |
|--------------|----------|
| "Cannot connect to Exchange Online" | Verify admin permissions |
| "No activities found" | Check date range and filters |
| "The term 'Connect-ExchangeOnline' is not recognized" | Run `Install-Module ExchangeOnlineManagement` |
| Throttling errors | Reduce `-ThrottleLimit` or increase `-IntervalMinutes` |

## Best Practices

1. **Regular Audits**: Schedule weekly/monthly executions
2. **Certificate Auth**: Recommended for automation
3. **Risk Prioritization**: Focus on "High Risk" activities first  
4. **Combine with Alerts**: Trigger workflows for sensitive activities
