<#
.SYNOPSIS
    Analyzes Azure storage accounts across all subscriptions to identify cost-saving opportunities, aligned with FinOps principles.
.DESCRIPTION
    Connects to Azure, retrieves storage account details, usage metrics, and costs, and generates a professional report with actionable recommendations and savings estimates. Outputs to console, CSV, Excel (with charts), and optionally PDF for client presentations.
.PARAMETER ExportCsv
    Export report to CSV.
.PARAMETER ExportExcel
    Export report to Excel with charts (requires ImportExcel module).
.PARAMETER ExportPdf
    Export a LaTeX template and optionally compile to PDF (if latexmk is available).
.PARAMETER ThresholdGB
    Low usage threshold in GB (default: 10).
.PARAMETER Days
    Analysis period in days (default: 30).
.PARAMETER OutputDir
    Output directory for logs and reports. Defaults to ./out under the current location.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [switch]$ExportExcel,

    [Parameter(Mandatory = $false)]
    [switch]$ExportPdf,

    [Parameter(Mandatory = $false)]
    [int]$ThresholdGB = 10,

    [Parameter(Mandatory = $false)]
    [int]$Days = 30,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = (Join-Path -Path (Get-Location) -ChildPath 'out')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure output directory exists
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# Import required modules
$requiredAzModules = @('Az.Accounts','Az.Storage','Az.Monitor','Az.Consumption')
foreach ($moduleName in $requiredAzModules) {
    try {
        Import-Module $moduleName -ErrorAction Stop
    }
    catch {
        Write-Warning "Module '$moduleName' is not available. Install with: Install-Module $moduleName -Scope CurrentUser"
        throw
    }
}

# Optional ImportExcel module for Excel export
$importExcelAvailable = $false
try {
    $null = Get-Module -ListAvailable -Name ImportExcel -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
    $importExcelAvailable = $true
}
catch {
    if ($ExportExcel) {
        Write-Warning "ImportExcel module not found. Excel export will be skipped. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }
}

# Pricing constants (2025 US East, approximate; adjust for region)
$pricing = @{
    HotLRS           = 0.0184   # $/GB/month
    CoolLRS          = 0.01     # $/GB/month
    ArchiveLRS       = 0.001    # $/GB/month
    CoolRetrieval    = 0.01     # $/GB for retrieval
    ArchiveRetrieval = 0.05     # $/GB for retrieval
    GRSMultiplier    = 2.0
    ZRSMultiplier    = 1.25
}

# Initialize logging
$timestampSuffix = Get-Date -Format 'yyyyMMdd_HHmm'
$logFile = Join-Path -Path $OutputDir -ChildPath "StorageCostAnalysis_$timestampSuffix.log"
function Write-Log {
    param ([Parameter(Mandatory)][string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
    Write-Verbose $Message
}

# Initialize report array
$report = New-Object System.Collections.Generic.List[object]
$scriptStartTime = Get-Date
Write-Host "[$($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))] Starting Azure Storage Cost Analysis" -ForegroundColor Green
Write-Log 'Script started.'

# Authenticate to Azure
try {
    if (-not (Get-AzContext)) {
        Connect-AzAccount -ErrorAction Stop | Out-Null
        Write-Log 'Authenticated to Azure successfully (new session).'
    }
    else {
        Write-Log 'Using existing Azure authentication context.'
    }
}
catch {
    Write-Host "Authentication failed: $_" -ForegroundColor Red
    Write-Log "Authentication failed: $_"
    exit 1
}

# Get all subscriptions
$subscriptions = @()
try {
    $subscriptions = Get-AzSubscription -ErrorAction Stop
}
catch {
    Write-Host "Failed to list subscriptions: $_" -ForegroundColor Red
    Write-Log "Failed to list subscriptions: $_"
    exit 1
}

$totalCurrentCost = 0.0
$totalPotentialSavings = 0.0

foreach ($subscription in $subscriptions) {
    Write-Host "===> Processing subscription: $($subscription.Name) ($($subscription.Id))" -ForegroundColor Yellow
    Write-Log "Processing subscription: $($subscription.Name) ($($subscription.Id))"

    try {
        Set-AzContext -Subscription $subscription.Id -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Failed to set context for subscription $($subscription.Name): $_" -ForegroundColor Red
        Write-Log "Failed to set context for subscription $($subscription.Name): $_"
        continue
    }

    # Get all storage accounts in the subscription
    $storageAccounts = @()
    try {
        $storageAccounts = Get-AzStorageAccount -ErrorAction Stop
    }
    catch {
        Write-Host "   |-- Failed to list storage accounts in subscription $($subscription.Name): $_" -ForegroundColor Red
        Write-Log "Failed to list storage accounts in subscription $($subscription.Name): $_"
        continue
    }

    if (-not $storageAccounts -or $storageAccounts.Count -eq 0) {
        Write-Host "   |-- No storage accounts found in subscription $($subscription.Name)" -ForegroundColor Yellow
        Write-Log "No storage accounts found in subscription $($subscription.Name)"
        continue
    }

    # Establish analysis window per subscription
    $metricWindowEnd   = Get-Date
    $metricWindowStart = $metricWindowEnd.AddDays(-[double]$Days)

    # Pre-fetch consumption usage details once per subscription for performance
    $costByResourceId = @{}
    try {
        $usageDetails = Get-AzConsumptionUsageDetail -StartDate $metricWindowStart -EndDate $metricWindowEnd -ErrorAction Stop
        $storageUsage  = $usageDetails | Where-Object { $_.ConsumedService -eq 'Microsoft.Storage' -and $_.ResourceId }
        foreach ($u in $storageUsage) {
            $rid = [string]$u.ResourceId
            if (-not $costByResourceId.ContainsKey($rid)) { $costByResourceId[$rid] = 0.0 }
            $costByResourceId[$rid] += [double]$u.PretaxCost
        }
    }
    catch {
        Write-Log "Failed to retrieve consumption usage details for subscription $($subscription.Name): $_"
    }

    foreach ($storageAccount in $storageAccounts) {
        Write-Host "   |-- Analyzing: $($storageAccount.StorageAccountName)" -ForegroundColor Cyan
        Write-Log "Analyzing storage account: $($storageAccount.StorageAccountName)"

        $accessTier = if ($null -ne $storageAccount.AccessTier -and "$($storageAccount.AccessTier)" -ne '') { [string]$storageAccount.AccessTier } else { 'N/A' }
        $redundancySkuName = [string]$storageAccount.Sku.Name
        $tagsJson = if ($null -ne $storageAccount.Tags -and $storageAccount.Tags.Keys.Count -gt 0) { $storageAccount.Tags | ConvertTo-Json -Compress } else { 'None' }

        $reportEntry = [PSCustomObject]@{
            Subscription       = $subscription.Name
            StorageAccount     = $storageAccount.StorageAccountName
            ResourceGroup      = $storageAccount.ResourceGroupName
            Location           = $storageAccount.Location
            AccessTier         = $accessTier
            Redundancy         = $redundancySkuName
            LifecyclePolicy    = 'Not Enabled'
            Tags               = $tagsJson
            UsedCapacityGB     = 0.0
            Transactions       = 0
            IngressGB          = 0.0
            EgressGB           = 0.0
            MonthlyCost        = 0.0
            Recommendations    = @()
            EstimatedSavings   = 0.0
            Implementation     = @()
        }

        # Lifecycle policy
        try {
            $policy = Get-AzStorageAccountManagementPolicy -ResourceGroupName $storageAccount.ResourceGroupName -StorageAccountName $storageAccount.StorageAccountName -ErrorAction Stop
            if ($null -ne $policy) { $reportEntry.LifecyclePolicy = 'Enabled' }
        }
        catch {
            Write-Log "Failed to retrieve lifecycle policy for $($storageAccount.StorageAccountName): $_"
        }

        # Usage metrics
        $resourceId = [string]$storageAccount.Id
        try {
            $capacityMetric = Get-AzMetric -ResourceId $resourceId -MetricName 'UsedCapacity' -TimeGrain 1.00:00:00 -StartTime $metricWindowStart -EndTime $metricWindowEnd -AggregationType Average -ErrorAction Stop
            $avgBytes = ($capacityMetric.Data | Where-Object { $_.Average -ne $null } | Measure-Object -Property Average -Average).Average
            if ($null -ne $avgBytes) { $reportEntry.UsedCapacityGB = [math]::Round(($avgBytes / 1GB), 2) }
        }
        catch {
            Write-Log "Failed to retrieve UsedCapacity for $($storageAccount.StorageAccountName): $_"
        }

        try {
            $transactionMetric = Get-AzMetric -ResourceId $resourceId -MetricName 'Transactions' -TimeGrain 1.00:00:00 -StartTime $metricWindowStart -EndTime $metricWindowEnd -AggregationType Total -ErrorAction Stop
            $reportEntry.Transactions = [int][math]::Round((($transactionMetric.Data | Where-Object { $_.Total -ne $null } | Measure-Object -Property Total -Sum).Sum), 0)
        }
        catch {
            Write-Log "Failed to retrieve Transactions for $($storageAccount.StorageAccountName): $_"
        }

        try {
            $ingressMetric = Get-AzMetric -ResourceId $resourceId -MetricName 'Ingress' -TimeGrain 1.00:00:00 -StartTime $metricWindowStart -EndTime $metricWindowEnd -AggregationType Total -ErrorAction Stop
            $totalIngressBytes = ($ingressMetric.Data | Where-Object { $_.Total -ne $null } | Measure-Object -Property Total -Sum).Sum
            if ($null -ne $totalIngressBytes) { $reportEntry.IngressGB = [math]::Round(($totalIngressBytes / 1GB), 2) }
        }
        catch {
            Write-Log "Failed to retrieve Ingress for $($storageAccount.StorageAccountName): $_"
        }

        try {
            $egressMetric = Get-AzMetric -ResourceId $resourceId -MetricName 'Egress' -TimeGrain 1.00:00:00 -StartTime $metricWindowStart -EndTime $metricWindowEnd -AggregationType Total -ErrorAction Stop
            $totalEgressBytes = ($egressMetric.Data | Where-Object { $_.Total -ne $null } | Measure-Object -Property Total -Sum).Sum
            if ($null -ne $totalEgressBytes) { $reportEntry.EgressGB = [math]::Round(($totalEgressBytes / 1GB), 2) }
        }
        catch {
            Write-Log "Failed to retrieve Egress for $($storageAccount.StorageAccountName): $_"
        }

        # Cost estimation (subscription-level cache)
        try {
            if ($costByResourceId.ContainsKey($resourceId)) {
                $reportEntry.MonthlyCost = [math]::Round([double]$costByResourceId[$resourceId], 2)
            }
            else {
                # Fallback by InstanceName when ResourceId is missing in usage (rare)
                $fallbackCost = 0.0
                try {
                    $ud = Get-AzConsumptionUsageDetail -StartDate $metricWindowStart -EndDate $metricWindowEnd -ErrorAction Stop |
                        Where-Object { $_.ConsumedService -eq 'Microsoft.Storage' -and $_.InstanceName -eq $storageAccount.StorageAccountName }
                    $fallbackCost = [double]([math]::Round((($ud | Measure-Object -Property PretaxCost -Sum).Sum), 2))
                }
                catch {}
                $reportEntry.MonthlyCost = $fallbackCost
            }
        }
        catch {
            Write-Log "Failed to retrieve cost data for $($storageAccount.StorageAccountName): $_"
        }

        # FinOps Recommendations
        $recommendations = New-Object System.Collections.Generic.List[string]
        $implementation  = New-Object System.Collections.Generic.List[string]
        $estimatedSavings = 0.0

        # Tag enforcement (safe null checks)
        $hasCostCenter = ($null -ne $storageAccount.Tags -and $storageAccount.Tags.ContainsKey('CostCenter'))
        $hasOwner      = ($null -ne $storageAccount.Tags -and $storageAccount.Tags.ContainsKey('Owner'))
        if (-not $hasCostCenter -or -not $hasOwner) {
            $recommendations.Add("Add 'CostCenter' and 'Owner' tags for cost attribution.")
            $implementation.Add("Set-AzStorageAccount -ResourceGroupName $($storageAccount.ResourceGroupName) -Name $($storageAccount.StorageAccountName) -Tag @{CostCenter='TBD'; Owner='TBD'}")
        }

        # Determine redundancy multiplier
        $skuUpper = $redundancySkuName.ToUpperInvariant()
        $redundancyMultiplier = if ($skuUpper -match 'GZRS|GRS') { $pricing.GRSMultiplier } elseif ($skuUpper -match 'ZRS') { $pricing.ZRSMultiplier } else { 1.0 }

        $transactionsPerDay = if ($Days -gt 0) { [math]::Round(($reportEntry.Transactions / [double]$Days), 2) } else { 0.0 }

        # Tier Optimization (basic heuristic)
        if ($reportEntry.AccessTier -eq 'Hot' -and $transactionsPerDay -lt 1000 -and $reportEntry.EgressGB -lt 1) {
            $currentCost   = $reportEntry.UsedCapacityGB * $pricing.HotLRS * $redundancyMultiplier
            $newCost       = $reportEntry.UsedCapacityGB * $pricing.CoolLRS * $redundancyMultiplier + ($reportEntry.EgressGB * $pricing.CoolRetrieval)
            $potentialSave = $currentCost - $newCost
            if ($potentialSave -gt 0) {
                $recommendations.Add("Move to Cool tier: Low transactions ($transactionsPerDay/day) and egress ($($reportEntry.EgressGB) GB). Saves ~$$([math]::Round($potentialSave,2))/month.")
                $implementation.Add("Set-AzStorageAccount -ResourceGroupName $($storageAccount.ResourceGroupName) -Name $($storageAccount.StorageAccountName) -AccessTier Cool")
                $estimatedSavings += $potentialSave
            }
        }
        elseif ($reportEntry.AccessTier -eq 'Cool' -and $transactionsPerDay -lt 100 -and $reportEntry.EgressGB -lt 0.1) {
            $currentCost   = $reportEntry.UsedCapacityGB * $pricing.CoolLRS * $redundancyMultiplier + ($reportEntry.EgressGB * $pricing.CoolRetrieval)
            $newCost       = $reportEntry.UsedCapacityGB * $pricing.ArchiveLRS * $redundancyMultiplier + ($reportEntry.EgressGB * $pricing.ArchiveRetrieval)
            $potentialSave = $currentCost - $newCost
            if ($potentialSave -gt 0) {
                $recommendations.Add("Move to Archive tier: Very low transactions ($transactionsPerDay/day) and egress ($($reportEntry.EgressGB) GB). Saves ~$$([math]::Round($potentialSave,2))/month.")
                $implementation.Add("Set-AzStorageAccount -ResourceGroupName $($storageAccount.ResourceGroupName) -Name $($storageAccount.StorageAccountName) -AccessTier Archive")
                $estimatedSavings += $potentialSave
            }
        }

        # Redundancy reduction (heuristic)
        $hasHighAvailabilityTag = ($null -ne $storageAccount.Tags -and $storageAccount.Tags.ContainsKey('HighAvailabilityRequired'))
        if ($skuUpper -match 'GZRS|GRS' -and -not $hasHighAvailabilityTag) {
            $currentCost   = $reportEntry.UsedCapacityGB * ($pricing.HotLRS * $pricing.GRSMultiplier)
            $newCost       = $reportEntry.UsedCapacityGB * $pricing.HotLRS
            $potentialSave = $currentCost - $newCost
            if ($potentialSave -gt 0) {
                $recommendations.Add("Switch to LRS: No high-availability tag. Saves ~$$([math]::Round($potentialSave,2))/month.")
                $implementation.Add("Set-AzStorageAccount -ResourceGroupName $($storageAccount.ResourceGroupName) -Name $($storageAccount.StorageAccountName) -SkuName Standard_LRS")
                $estimatedSavings += $potentialSave
            }
        }

        # Lifecycle policy suggestion (assume 50% eligible)
        if ($reportEntry.LifecyclePolicy -eq 'Not Enabled') {
            $potentialSave = $reportEntry.UsedCapacityGB * ($pricing.HotLRS - $pricing.CoolLRS) * 0.5 * $redundancyMultiplier
            $recommendations.Add("Enable lifecycle policy: Auto-tier blobs >30 days to Cool, >90 days to Archive. Saves ~$$([math]::Round($potentialSave,2))/month.")
            $implementation.Add("# Example (adjust to your policy cmdlets and az module version):")
            $implementation.Add("# Set-AzStorageAccountManagementPolicy with rules to tier >30d to Cool and >90d to Archive")
            $estimatedSavings += $potentialSave
        }

        # Deletion/Archival suggestion for underutilized accounts
        if ($reportEntry.UsedCapacityGB -lt $ThresholdGB -and $transactionsPerDay -lt 100) {
            $recommendations.Add("Delete/archive: Low capacity ($($reportEntry.UsedCapacityGB) GB) and transactions ($transactionsPerDay/day). Saves ~$$([math]::Round($reportEntry.MonthlyCost,2))/month.")
            $implementation.Add("Remove-AzStorageAccount -ResourceGroupName $($storageAccount.ResourceGroupName) -Name $($storageAccount.StorageAccountName) -Force")
            $estimatedSavings += $reportEntry.MonthlyCost
        }

        # Reservations suggestion
        if ($reportEntry.UsedCapacityGB -gt 100 -and $reportEntry.MonthlyCost -gt 100) {
            $potentialSave = $reportEntry.MonthlyCost * 0.3  # Assume 30% savings
            $recommendations.Add("Use reserved capacity: High capacity ($($reportEntry.UsedCapacityGB) GB) and cost ($$([math]::Round($reportEntry.MonthlyCost,2))). Saves ~$$([math]::Round($potentialSave,2))/month.")
            $implementation.Add("Visit Azure Portal > Reservations to purchase reserved capacity.")
            $estimatedSavings += $potentialSave
        }

        $reportEntry.Recommendations = ($recommendations -join '; ')
        $reportEntry.Implementation  = ($implementation -join '; ')
        $reportEntry.EstimatedSavings = [math]::Round($estimatedSavings, 2)

        $totalCurrentCost     += $reportEntry.MonthlyCost
        $totalPotentialSavings += $reportEntry.EstimatedSavings

        $null = $report.Add($reportEntry)
    }
}

# Summary table
$summary = [PSCustomObject]@{
    TotalAccounts         = $report.Count
    TotalCurrentCost      = [math]::Round($totalCurrentCost, 2)
    TotalPotentialSavings = [math]::Round($totalPotentialSavings, 2)
    SavingsPercentage     = if ($totalCurrentCost -gt 0) { [math]::Round(($totalPotentialSavings / $totalCurrentCost) * 100, 1) } else { 0.0 }
}

Write-Host "`n=== Cost Optimization Summary ===" -ForegroundColor Cyan
Write-Host "Total Accounts: $($summary.TotalAccounts)" -ForegroundColor Cyan
Write-Host "Current Cost: $$($summary.TotalCurrentCost)/month" -ForegroundColor Cyan
Write-Host "Potential Savings: $$($summary.TotalPotentialSavings)/month ($($summary.SavingsPercentage)%)" -ForegroundColor Green
Write-Log "Summary: Total Accounts=$($summary.TotalAccounts), Current Cost=$$($summary.TotalCurrentCost), Savings=$$($summary.TotalPotentialSavings) ($($summary.SavingsPercentage)%)"

# Console detailed output with color-coding
Write-Host "`n=== Detailed Report ===" -ForegroundColor Cyan
foreach ($row in $report) {
    $color = if ($row.MonthlyCost -gt 100) { 'Red' } elseif ($row.UsedCapacityGB -lt $ThresholdGB) { 'Green' } else { 'White' }
    Write-Host ("$($row.StorageAccount): $$($row.MonthlyCost)/month, $($row.UsedCapacityGB) GB, $($row.Recommendations)") -ForegroundColor $color
}

# Export to CSV
if ($ExportCsv) {
    $csvFile = Join-Path -Path $OutputDir -ChildPath "StorageCostReport_$timestampSuffix.csv"
    $report | Select-Object Subscription, StorageAccount, ResourceGroup, Location, AccessTier, Redundancy, LifecyclePolicy, Tags, UsedCapacityGB, Transactions, IngressGB, EgressGB, MonthlyCost, Recommendations, EstimatedSavings, Implementation |
        Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    Write-Host "CSV report exported to $csvFile" -ForegroundColor Green
    Write-Log "CSV report exported to $csvFile"
}

# Excel export with charts (if ImportExcel is available)
if ($ExportExcel -and $importExcelAvailable) {
    try {
        $excelFile = Join-Path -Path $OutputDir -ChildPath "StorageCostReport_$timestampSuffix.xlsx"

        $chart1 = New-ExcelChartDefinition -Title 'Cost by Access Tier' -ChartType Pie -XRange 'Storage Accounts[AccessTier]' -YRange 'Storage Accounts[MonthlyCost]'
        $chart2 = New-ExcelChartDefinition -Title 'Potential Savings by Account' -ChartType ColumnClustered -XRange 'Storage Accounts[StorageAccount]' -YRange 'Storage Accounts[EstimatedSavings]'

        $report | Export-Excel -Path $excelFile -AutoSize -TableName 'StorageCostAnalysis' -TableStyle Light8 -WorksheetName 'Storage Accounts' -Title 'Azure Storage Cost Optimization Report' -ExcelChartDefinition @($chart1,$chart2) -ClearSheet

        Write-Host "Excel report with charts exported to $excelFile" -ForegroundColor Green
        Write-Log "Excel report with charts exported to $excelFile"
    }
    catch {
        Write-Warning "Failed to export Excel: $_"
        Write-Log "Failed to export Excel: $_"
    }
}

# PDF report using LaTeX (template; optionally compile with latexmk if available)
if ($ExportPdf) {
    $latexContent = @"
\documentclass{article}
\usepackage{booktabs}
\usepackage{geometry}
\usepackage{xcolor}
\geometry{a4paper, margin=1in}

\begin{document}

\title{Azure Storage Cost Optimization Report}
\author{FinOps Analysis}
\date{$(Get-Date -Format 'MMMM dd, yyyy')}
\maketitle

\section*{Summary}
This report analyzes $($summary.TotalAccounts) storage accounts across all Azure subscriptions. Current monthly cost is \$\textbf{$($summary.TotalCurrentCost)}, with potential savings of \$\textbf{$($summary.TotalPotentialSavings)} (\textbf{$($summary.SavingsPercentage)}\%) by implementing recommendations. Pricing is based on 2025 US East rates; verify with the Azure Pricing Calculator for your region.

\begin{table}[h]
\centering
\begin{tabular}{lr}
\toprule
Metric & Value \\
\midrule
Total Accounts & $($summary.TotalAccounts) \\
Current Cost (\$/month) & $($summary.TotalCurrentCost) \\
Potential Savings (\$/month) & $($summary.TotalPotentialSavings) \\
Savings Percentage & $($summary.SavingsPercentage)\% \\
\bottomrule
\end{tabular}
\caption{Summary of Findings}
\end{table}

\section*{Recommendations}
\begin{itemize}
    \item \textbf{Tier Optimization}: Move low-access accounts to Cool/Archive tiers.
    \item \textbf{Redundancy Reduction}: Switch to LRS where high availability is not required.
    \item \textbf{Lifecycle Policies}: Enable auto-tiering for blobs >30 days to Cool, >90 days to Archive.
    \item \textbf{Deletion/Archival}: Remove underutilized accounts.
    \item \textbf{Reservations}: Use reserved capacity for high-usage accounts.
    \item \textbf{Tagging}: Add CostCenter/Owner tags for accountability.
\end{itemize}

\section*{Next Steps}
Run the provided PowerShell commands in the CSV/Excel report's ``Implementation'' column to apply changes. Contact your Azure administrator for assistance or use the Azure Portal.

\end{document}
"@

    try {
        $latexFile = Join-Path -Path $OutputDir -ChildPath "StorageCostReport_$timestampSuffix.tex"
        $latexContent | Out-File -FilePath $latexFile -Encoding UTF8
        Write-Host "PDF report template exported to $latexFile" -ForegroundColor Green
        Write-Log "PDF report template exported to $latexFile"

        # Attempt to compile with latexmk if available
        $latexmk = Get-Command latexmk -ErrorAction SilentlyContinue
        if ($latexmk) {
            try {
                $latexDir = Split-Path -Path $latexFile -Parent
                Push-Location $latexDir
                & $latexmk -pdf -quiet (Split-Path -Leaf $latexFile) | Out-Null
                Pop-Location
                $pdfFile = [System.IO.Path]::ChangeExtension($latexFile,'pdf')
                if (Test-Path -LiteralPath $pdfFile) {
                    Write-Host "PDF compiled: $pdfFile" -ForegroundColor Green
                    Write-Log "PDF compiled: $pdfFile"
                }
            }
            catch {
                Write-Warning "Failed to compile PDF with latexmk: $_"
                Write-Log "Failed to compile PDF with latexmk: $_"
            }
        }
        else {
            Write-Host 'latexmk not found; to compile the .tex file to PDF, install TeX Live or MikTeX and run: latexmk -pdf <file.tex>' -ForegroundColor DarkYellow
        }
    }
    catch {
        Write-Warning "Failed to write LaTeX template: $_"
        Write-Log "Failed to write LaTeX template: $_"
    }
}

# End script
$scriptEndTime = Get-Date
Write-Host "[$($scriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))] Script finished." -ForegroundColor Green
Write-Host "Total minutes elapsed: $('{0:N2}' -f ($scriptEndTime - $scriptStartTime).TotalMinutes)" -ForegroundColor DarkYellow
Write-Log "Script finished. Total minutes elapsed: $('{0:N2}' -f ($scriptEndTime - $scriptStartTime).TotalMinutes)"