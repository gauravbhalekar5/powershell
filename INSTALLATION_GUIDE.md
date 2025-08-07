# Azure Premium Disk IOPS Analysis Script - Installation Guide

## Prerequisites

### 1. PowerShell Version
- **Windows PowerShell 5.1** or **PowerShell 7.x**
- Compatible with **PowerShell ISE** and **PowerShell Console**

### 2. Azure PowerShell Modules
Install the required Azure PowerShell modules:

```powershell
# Install all required modules at once
Install-Module -Name Az.Accounts, Az.Compute, Az.Storage, Az.Resources, Az.Monitor -Force -AllowClobber

# Or install the complete Az module (includes all sub-modules)
Install-Module -Name Az -Force -AllowClobber
```

### 3. Optional: Excel Export Module
For enhanced Excel reporting with multiple worksheets:

```powershell
Install-Module -Name ImportExcel -Force -AllowClobber
```

## Azure Permissions Required

### Minimum Role Requirements
Your Azure account needs **Reader** role on:
- Target subscriptions
- Resource groups containing premium disks
- Virtual machines (for power state and attachment info)

### Recommended Permissions
- **Reader** role at subscription level
- **Monitoring Reader** role (for Azure Monitor metrics)

## Installation Steps

### Step 1: Download the Script
Save the `Azure-Premium-Disk-IOPS-Analysis.ps1` file to your local machine.

### Step 2: Set Execution Policy (if needed)
```powershell
# Check current execution policy
Get-ExecutionPolicy

# If restricted, set to allow script execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Step 3: Test Azure Connection
```powershell
# Connect to Azure
Connect-AzAccount

# Verify access to target subscriptions
Get-AzSubscription | Where-Object { $_.Name -like "*DENV*" -or $_.Name -like "*Daikin*" }
```

## Usage Examples

### Basic Usage
```powershell
# Run with default settings (7-day analysis, 30% threshold)
.\Azure-Premium-Disk-IOPS-Analysis.ps1
```

### Custom Configuration
```powershell
# Custom analysis period and threshold
.\Azure-Premium-Disk-IOPS-Analysis.ps1 -AnalysisDays 14 -UtilizationThreshold 25

# Specific subscriptions only
.\Azure-Premium-Disk-IOPS-Analysis.ps1 -TargetSubscriptions @("DENV Prod", "DENV non-prod")

# CSV export only (skip Excel)
.\Azure-Premium-Disk-IOPS-Analysis.ps1 -ExportToExcel:$false

# Custom output location
.\Azure-Premium-Disk-IOPS-Analysis.ps1 -OutputPath "C:\Reports\DiskAnalysis_$(Get-Date -Format 'yyyyMMdd')"
```

### Advanced Usage
```powershell
# Complete analysis with all options
.\Azure-Premium-Disk-IOPS-Analysis.ps1 `
    -TargetSubscriptions @("DENV Prod", "DENV non-prod", "Daikin Europe") `
    -AnalysisDays 30 `
    -UtilizationThreshold 20 `
    -OutputPath "C:\DiskAnalysis\Monthly_Report" `
    -ExportToExcel:$true `
    -DetailedReport:$true
```

## Output Files

The script generates three types of output files:

### 1. Excel Workbook (if ImportExcel module available)
- **File**: `[OutputPath].xlsx`
- **Sheets**:
  - `DiskAnalysis`: Complete disk details and recommendations
  - `Summary`: Aggregated statistics by recommendation type
  - `BySubscription`: Per-subscription breakdown

### 2. CSV Export
- **File**: `[OutputPath].csv`
- **Content**: Complete raw data for further analysis

### 3. Text Report
- **File**: `[OutputPath].txt`
- **Content**: Executive summary with business justification

## Troubleshooting

### Common Issues

#### 1. Module Not Found
```
Error: Az.Accounts module not found
```
**Solution:**
```powershell
Install-Module -Name Az -Force -AllowClobber
```

#### 2. Access Denied
```
Error: Failed to get target subscriptions
```
**Solution:**
- Verify Azure login: `Get-AzContext`
- Check subscription access: `Get-AzSubscription`
- Contact Azure admin for Reader permissions

#### 3. No Metrics Available
```
Warning: Could not retrieve metrics for disk [DiskName]
```
**Explanation:** This is normal for:
- Unattached disks
- Stopped VMs
- Disks without recent activity

The script uses conservative estimates in these cases.

#### 4. Excel Export Failed
```
Warning: ImportExcel module not available
```
**Solution:**
```powershell
Install-Module -Name ImportExcel -Force
```
Or use CSV export: `-ExportToExcel:$false`

### Performance Considerations

#### Large Environments
For environments with many disks (500+):
- Consider running during off-peak hours
- Use longer analysis periods (-AnalysisDays 30) for more accurate data
- Process subscriptions individually if timeout occurs

#### Metric Collection
- Azure Monitor metrics may take 5-15 minutes to become available
- For immediate analysis of newly created disks, the script uses conservative estimates
- Historical data (7+ days) provides most accurate recommendations

## Best Practices

### 1. Regular Analysis
- Run monthly for ongoing optimization
- Compare results over time to track improvements
- Schedule during maintenance windows for minimal impact

### 2. Validation Process
- Review unattached disk recommendations with application owners
- Validate migration candidates in non-production first
- Monitor performance after migration

### 3. Documentation
- Keep analysis reports for audit purposes
- Document migration decisions and outcomes
- Share cost savings with stakeholders

### 4. Automation
Consider scheduling the script:

```powershell
# Example: Windows Task Scheduler PowerShell command
powershell.exe -File "C:\Scripts\Azure-Premium-Disk-IOPS-Analysis.ps1" -OutputPath "C:\Reports\Monthly_$(Get-Date -Format 'yyyyMM')"
```

## Support and Updates

### Script Updates
- Check for updated Azure disk specifications
- Review pricing information quarterly
- Update target subscriptions as environment changes

### Getting Help
1. Review the detailed error messages and troubleshooting steps
2. Check Azure PowerShell module versions: `Get-Module Az.*`
3. Verify Azure permissions and subscription access
4. Consult Azure documentation for latest disk specifications

## Security Notes

- Script requires read-only access to Azure resources
- No data is modified or deleted by the script
- Output files may contain sensitive subscription and resource information
- Store output files securely and limit access appropriately