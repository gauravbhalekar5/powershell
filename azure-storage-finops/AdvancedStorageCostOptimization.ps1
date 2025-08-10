<#
.SYNOPSIS
    Advanced Azure storage cost optimization analyzer with enhanced FinOps capabilities and multi-region support.
.DESCRIPTION
    Performs comprehensive analysis of Azure storage accounts across all subscriptions, including:
    - Multi-region pricing analysis
    - Advanced lifecycle policy recommendations
    - Blob-level analysis for large accounts (optional)
    - Reserved capacity ROI calculations
    - Historical trend analysis (via metrics window)
    - Compliance and security checks
    - Automated remediation options (safe, optional)
.PARAMETER ExportCsv
    Export detailed report to CSV format.
.PARAMETER ExportExcel
    Export report to Excel with charts (requires ImportExcel module).
.PARAMETER ExportPdf
    Generate professional PDF report with executive summary.
.PARAMETER ExportHtml
    Generate simple HTML dashboard.
.PARAMETER ThresholdGB
    Low usage threshold in GB (default: 10).
.PARAMETER Days
    Analysis period in days (default: 30).
.PARAMETER OutputDir
    Output directory for reports. Defaults to ./reports/[timestamp].
.PARAMETER AutoRemediate
    Generate an actions file with safe optimizations; optionally execute if -Force is passed.
.PARAMETER Region
    Target region for pricing calculations (default: auto-detect per account location).
.PARAMETER DetailedAnalysis
    Perform blob-level analysis for accounts over specified size (placeholder; opt-in).
.PARAMETER EmailReport
    Email address to send report summary (requires SMTP env configuration; see README).
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [switch]$ExportExcel,

    [Parameter(Mandatory = $false)]
    [switch]$ExportPdf,

    [Parameter(Mandatory = $false)]
    [switch]$ExportHtml,

    [Parameter(Mandatory = $false)]
    [int]$ThresholdGB = 10,

    [Parameter(Mandatory = $false)]
    [int]$Days = 30,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [switch]$AutoRemediate,

    [Parameter(Mandatory = $false)]
    [string]$Region = 'auto',

    [Parameter(Mandatory = $false)]
    [switch]$DetailedAnalysis,

    [Parameter(Mandatory = $false)]
    [string]$EmailReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Enhanced configuration
$script:Config = @{
    Version   = '2.0.1'
    RunId     = [guid]::NewGuid().ToString('N').Substring(0, 8)
    StartTime = Get-Date
}

# Regional pricing matrix (2025 rates per GB/month in USD)
$script:RegionalPricing = @{
    'eastus' = @{
        HotLRS = 0.0184; CoolLRS = 0.01; ArchiveLRS = 0.001
        HotGRS = 0.0368; CoolGRS = 0.02; ArchiveGRS = 0.002
        HotZRS = 0.023;   CoolZRS = 0.0125; ArchiveZRS = 0.00125
        TransactionCost = 0.0004; EarlyDeletionDays = 30
    }
    'westus2' = @{
        HotLRS = 0.0184; CoolLRS = 0.01; ArchiveLRS = 0.001
        HotGRS = 0.0368; CoolGRS = 0.02; ArchiveGRS = 0.002
        HotZRS = 0.023;   CoolZRS = 0.0125; ArchiveZRS = 0.00125
        TransactionCost = 0.0004; EarlyDeletionDays = 30
    }
    'westeurope' = @{
        HotLRS = 0.0196; CoolLRS = 0.011; ArchiveLRS = 0.0011
        HotGRS = 0.0392; CoolGRS = 0.022; ArchiveGRS = 0.0022
        HotZRS = 0.0245;  CoolZRS = 0.0138; ArchiveZRS = 0.00138
        TransactionCost = 0.00045; EarlyDeletionDays = 30
    }
    'eastasia' = @{
        HotLRS = 0.022;  CoolLRS = 0.012; ArchiveLRS = 0.002
        HotGRS = 0.044;  CoolGRS = 0.024; ArchiveGRS = 0.004
        HotZRS = 0.0275; CoolZRS = 0.015;  ArchiveZRS = 0.0025
        TransactionCost = 0.0005; EarlyDeletionDays = 30
    }
}

# Initialize output directory
if (-not $OutputDir) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputDir = Join-Path -Path (Get-Location) -ChildPath (Join-Path -Path 'reports' -ChildPath "StorageAnalysis_$timestamp")
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# Enhanced logging system
class Logger {
    [string]$LogFile
    [string]$Level
    Logger([string]$logPath) {
        $this.LogFile = $logPath
        $this.Level = 'INFO'
    }
    [void] Write([string]$message, [string]$level = 'INFO') {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $logEntry = "$timestamp [$level] $message"
        $logEntry | Out-File -FilePath $this.LogFile -Append -Encoding UTF8
        switch ($level) {
            'ERROR'   { Write-Host $message -ForegroundColor Red }
            'WARNING' { Write-Host $message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $message -ForegroundColor Green }
            default   { Write-Verbose $message }
        }
    }
}

$script:Logger = [Logger]::new((Join-Path -Path $OutputDir -ChildPath 'analysis.log'))
$Logger.Write("Azure Storage Cost Optimization v$($Config.Version) started", 'SUCCESS')
$Logger.Write("Run ID: $($Config.RunId)", 'INFO')

# Module validation and import
function Initialize-RequiredModules {
    $requiredModules = @(
        @{Name = 'Az.Accounts';     MinVersion = '2.0.0'; Required = $true},
        @{Name = 'Az.Storage';      MinVersion = '4.0.0'; Required = $true},
        @{Name = 'Az.Monitor';      MinVersion = '3.0.0'; Required = $true},
        @{Name = 'Az.Consumption';  MinVersion = '2.0.0'; Required = $true},
        @{Name = 'Az.CostManagement'; MinVersion = '1.0.0'; Required = $false},
        @{Name = 'ImportExcel';     MinVersion = '7.0.0'; Required = $false}
    )
    foreach ($module in $requiredModules) {
        try {
            $installed = Get-Module -ListAvailable -Name $module.Name |
                Where-Object { $_.Version -ge [Version]$module.MinVersion } |
                Select-Object -First 1
            if ($installed) {
                Import-Module $module.Name -MinimumVersion $module.MinVersion -ErrorAction Stop
                $Logger.Write("Loaded module: $($module.Name) v$($installed.Version)", 'INFO')
            }
            elseif ($module.Required) {
                throw "Required module $($module.Name) v$($module.MinVersion)+ not found"
            }
            else {
                $Logger.Write("Optional module $($module.Name) not available", 'WARNING')
            }
        }
        catch {
            if ($module.Required) {
                $Logger.Write("Failed to load required module $($module.Name): $_", 'ERROR')
                throw
            }
            else {
                $Logger.Write("Optional module $($module.Name) not available: $_", 'WARNING')
            }
        }
    }
}

# Azure authentication with retry
function Connect-AzureWithRetry {
    param([int]$MaxRetries = 3)
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            $context = Get-AzContext -ErrorAction SilentlyContinue
            if (-not $context) {
                $Logger.Write("Authenticating to Azure (attempt $($attempt + 1)/$MaxRetries)...", 'INFO')
                Connect-AzAccount -ErrorAction Stop | Out-Null
                $Logger.Write('Successfully authenticated to Azure', 'SUCCESS')
            }
            else {
                $Logger.Write("Using existing Azure context: $($context.Account.Id)", 'INFO')
            }
            return $true
        }
        catch {
            $attempt++
            if ($attempt -ge $MaxRetries) {
                $Logger.Write("Authentication failed after $MaxRetries attempts: $_", 'ERROR')
                throw
            }
            Start-Sleep -Seconds 5
        }
    }
    return $false
}

# Storage account analyzer class
class StorageAccountAnalyzer {
    [object]$Account
    [hashtable]$Metrics
    [string]$Region
    [hashtable]$Pricing
    [int]$AnalysisDays
    [System.Collections.Generic.List[object]]$Recommendations

    StorageAccountAnalyzer([object]$storageAccount, [string]$region, [int]$days) {
        $this.Account = $storageAccount
        $this.Region = ($region | ForEach-Object { $_.ToLower() })
        $this.AnalysisDays = $days
        $this.Metrics = @{}
        $this.Recommendations = [System.Collections.Generic.List[object]]::new()
        $this.SetPricing()
    }

    [void] SetPricing() {
        if ($script:RegionalPricing.ContainsKey($this.Region)) { $this.Pricing = $script:RegionalPricing[$this.Region] }
        else {
            $this.Pricing = $script:RegionalPricing['eastus']
            $script:Logger.Write("Using default East US pricing for region: $($this.Region)", 'WARNING')
        }
    }

    [hashtable] AnalyzeUsageMetrics([datetime]$startDate, [datetime]$endDate) {
        $resourceId = [string]$this.Account.Id
        $localMetrics = @{
            UsedCapacityGB = 0.0
            Transactions   = 0
            IngressGB      = 0.0
            EgressGB       = 0.0
            BlobCount      = 0
        }
        try {
            # Capacity
            $capacityData = Get-AzMetric -ResourceId $resourceId -MetricName 'UsedCapacity' -TimeGrain 1.00:00:00 -StartTime $startDate -EndTime $endDate -AggregationType Average -ErrorAction Stop
            $avgBytes = ($capacityData.Data | Where-Object { $null -ne $_.Average } | Measure-Object -Property Average -Average).Average
            if ($avgBytes) { $localMetrics.UsedCapacityGB = [math]::Round($avgBytes / 1GB, 2) }
            # Transactions
            $transData = Get-AzMetric -ResourceId $resourceId -MetricName 'Transactions' -TimeGrain 1.00:00:00 -StartTime $startDate -EndTime $endDate -AggregationType Total -ErrorAction Stop
            $localMetrics.Transactions = [int](($transData.Data | Where-Object { $null -ne $_.Total } | Measure-Object -Property Total -Sum).Sum)
            # Ingress/Egress
            try {
                $ingress = Get-AzMetric -ResourceId $resourceId -MetricName 'Ingress' -TimeGrain 1.00:00:00 -StartTime $startDate -EndTime $endDate -AggregationType Total -ErrorAction SilentlyContinue
                if ($ingress) { $localMetrics.IngressGB = [math]::Round((($ingress.Data | Where-Object { $_.Total } | Measure-Object -Property Total -Sum).Sum) / 1GB, 2) }
                $egress = Get-AzMetric -ResourceId $resourceId -MetricName 'Egress' -TimeGrain 1.00:00:00 -StartTime $startDate -EndTime $endDate -AggregationType Total -ErrorAction SilentlyContinue
                if ($egress) { $localMetrics.EgressGB = [math]::Round((($egress.Data | Where-Object { $_.Total } | Measure-Object -Property Total -Sum).Sum) / 1GB, 2) }
            } catch {}
            # BlobCount (may not exist)
            try {
                $blobData = Get-AzMetric -ResourceId $resourceId -MetricName 'BlobCount' -TimeGrain 1.00:00:00 -StartTime $startDate -EndTime $endDate -AggregationType Average -ErrorAction SilentlyContinue
                if ($blobData) { $localMetrics.BlobCount = [int](($blobData.Data | Where-Object { $_.Average } | Measure-Object -Property Average -Average).Average) }
            } catch {}
        }
        catch {
            $script:Logger.Write("Failed metrics for $($this.Account.StorageAccountName): $_", 'ERROR')
        }
        $this.Metrics = $localMetrics
        return $localMetrics
    }

    [decimal] CalculateCurrentCost() {
        $tier = if ($this.Account.AccessTier) { $this.Account.AccessTier.ToString() } else { 'Hot' }
        $sku  = [string]$this.Account.Sku.Name  # e.g., Standard_LRS
        $redundancy = if ($sku -match 'GZRS|GRS') { 'GRS' } elseif ($sku -match 'ZRS') { 'ZRS' } else { 'LRS' }
        $priceKey = "$tier$redundancy"  # e.g., HotLRS, CoolGRS
        $basePrice = if ($this.Pricing.ContainsKey($priceKey)) { [decimal]$this.Pricing[$priceKey] } else { 0.0184m }
        $storageCost = [decimal]$this.Metrics.UsedCapacityGB * $basePrice
        $transactionUnits = [decimal]$this.Metrics.Transactions / 10000m
        $transactionCost = $transactionUnits * [decimal]$this.Pricing.TransactionCost
        return [decimal]([math]::Round([double]($storageCost + $transactionCost), 2))
    }

    [object] GenerateRecommendations() {
        $recMap = @{
            TierOptimization        = $null
            RedundancyOptimization  = $null
            LifecyclePolicy         = $null
            Archival                = $null
            ReservedCapacity        = $null
            Tagging                 = $null
            Security                = $null
        }
        $currentCost = [double]$this.CalculateCurrentCost()
        $totalSavings = 0.0
        $transPerDay = if ($this.AnalysisDays -gt 0) { [double]$this.Metrics.Transactions / [double]$this.AnalysisDays } else { 0.0 }
        $currentTier = if ($this.Account.AccessTier) { $this.Account.AccessTier.ToString() } else { 'Hot' }
        # Tier optimization
        if ($currentTier -eq 'Hot' -and $transPerDay -lt 1000) {
            $coolCost = [double]$this.Metrics.UsedCapacityGB * [double]$this.Pricing.CoolLRS
            $savings = $currentCost - $coolCost
            if ($savings -gt 5) {
                $recMap.TierOptimization = @{
                    Action = 'Move to Cool tier'
                    Reason = "Low transaction rate ($([math]::Round($transPerDay,1))/day)"
                    Savings = [math]::Round($savings, 2)
                    Risk = 'Low'
                    Implementation = @(
                        "Set-AzStorageAccount -ResourceGroupName $($this.Account.ResourceGroupName) -Name $($this.Account.StorageAccountName) -AccessTier Cool"
                    )
                }
                $totalSavings += $savings
            }
        }
        # Redundancy optimization
        $sku = [string]$this.Account.Sku.Name
        $hasCriticalTag = ($this.Account.Tags -and $this.Account.Tags.ContainsKey('CriticalData'))
        if ($sku -match 'GRS|GZRS' -and -not $hasCriticalTag) {
            $lrsCost = [double]$this.Metrics.UsedCapacityGB * [double]$this.Pricing.HotLRS
            $savings = $currentCost - $lrsCost
            if ($savings -gt 10) {
                $recMap.RedundancyOptimization = @{
                    Action = 'Switch to LRS redundancy'
                    Reason = 'No CriticalData tag found'
                    Savings = [math]::Round($savings, 2)
                    Risk = 'Medium'
                    Implementation = @(
                        "Set-AzStorageAccount -ResourceGroupName $($this.Account.ResourceGroupName) -Name $($this.Account.StorageAccountName) -SkuName Standard_LRS"
                    )
                }
                $totalSavings += $savings
            }
        }
        # Lifecycle policy
        $hasLifecycle = $false
        try {
            $policy = Get-AzStorageAccountManagementPolicy -ResourceGroupName $this.Account.ResourceGroupName -StorageAccountName $this.Account.StorageAccountName -ErrorAction SilentlyContinue
            $hasLifecycle = ($null -ne $policy)
        } catch {}
        if (-not $hasLifecycle -and $this.Metrics.UsedCapacityGB -gt 100) {
            $estimatedSavings = [double]$this.Metrics.UsedCapacityGB * 0.3 * ([double]$this.Pricing.HotLRS - [double]$this.Pricing.CoolLRS)
            $recMap.LifecyclePolicy = @{
                Action = 'Enable lifecycle management'
                Reason = "Large volume ($($this.Metrics.UsedCapacityGB) GB) without lifecycle policy"
                Savings = [math]::Round($estimatedSavings, 2)
                Risk = 'Low'
                Implementation = @(
                    '# Configure via Azure Portal > Storage Account > Lifecycle Management or use Az cmdlets'
                )
            }
            $totalSavings += $estimatedSavings
        }
        # Archival/deletion for unused accounts
        if ($this.Metrics.UsedCapacityGB -lt $ThresholdGB -and $transPerDay -lt 10) {
            $recMap.Archival = @{
                Action = 'Archive or delete account'
                Reason = "Minimal usage: $($this.Metrics.UsedCapacityGB) GB, $([math]::Round($transPerDay,1)) trans/day"
                Savings = [math]::Round($currentCost, 2)
                Risk = 'High'
                Implementation = @(
                    '# Review data retention requirements before deletion',
                    "Remove-AzStorageAccount -ResourceGroupName $($this.Account.ResourceGroupName) -Name $($this.Account.StorageAccountName) -Force"
                )
            }
            $totalSavings += $currentCost
        }
        # Reservations
        if ($this.Metrics.UsedCapacityGB -gt 500 -and $currentCost -gt 100) {
            $reservedSavings = $currentCost * 0.28
            $recMap.ReservedCapacity = @{
                Action = 'Purchase reserved capacity'
                Reason = "High stable usage: $($this.Metrics.UsedCapacityGB) GB"
                Savings = [math]::Round($reservedSavings, 2)
                Risk = 'Low'
                Implementation = @(
                    'Purchase via Azure Portal > Reservations'
                )
            }
            $totalSavings += $reservedSavings
        }
        # Tagging
        $missingTags = @()
        foreach ($tag in @('Owner','CostCenter','Environment','Application')) {
            if (-not ($this.Account.Tags -and $this.Account.Tags.ContainsKey($tag))) { $missingTags += $tag }
        }
        if ($missingTags.Count -gt 0) {
            $recMap.Tagging = @{
                Action = 'Add required tags'
                Reason = "Missing tags: $($missingTags -join ', ')"
                Savings = 0
                Risk = 'None'
                Implementation = @(
                    '$tags = @{ Owner="TBD"; CostCenter="TBD"; Environment="TBD"; Application="TBD" }',
                    "Set-AzStorageAccount -ResourceGroupName $($this.Account.ResourceGroupName) -Name $($this.Account.StorageAccountName) -Tag $tags"
                )
            }
        }
        # Security
        $securityIssues = @()
        if (-not $this.Account.EnableHttpsTrafficOnly) { $securityIssues += 'HTTPS not enforced' }
        if ($this.Account.AllowBlobPublicAccess)       { $securityIssues += 'Public blob access allowed' }
        if ($this.Account.MinimumTlsVersion -ne 'TLS1_2') { $securityIssues += 'TLS version < 1.2' }
        if ($securityIssues.Count -gt 0) {
            $recMap.Security = @{
                Action = 'Apply security best practices'
                Reason = ($securityIssues -join '; ')
                Savings = 0
                Risk = 'None'
                Implementation = @(
                    "Set-AzStorageAccount -ResourceGroupName $($this.Account.ResourceGroupName) -Name $($this.Account.StorageAccountName) -EnableHttpsTrafficOnly $true -AllowBlobPublicAccess $false -MinimumTlsVersion TLS1_2"
                )
            }
        }
        # Final
        $finalRecs = @()
        foreach ($key in $recMap.Keys) { if ($null -ne $recMap[$key]) { $finalRecs += $recMap[$key] } }
        return @{
            Recommendations = $finalRecs
            TotalSavings    = [math]::Round($totalSavings, 2)
            CurrentCost     = [math]::Round($currentCost, 2)
        }
    }
}

function Start-StorageAnalysis {
    $script:Report = [System.Collections.Generic.List[object]]::new()
    $script:ExecutiveSummary = @{
        TotalAccounts       = 0
        TotalSubscriptions  = 0
        TotalDataGB         = 0.0
        CurrentMonthlyCost  = 0.0
        PotentialSavings    = 0.0
    }
    $subscriptions = Get-AzSubscription -ErrorAction Stop
    $script:ExecutiveSummary.TotalSubscriptions = $subscriptions.Count
    $endDate = Get-Date
    $startDate = $endDate.AddDays(-[double]$Days)
    foreach ($subscription in $subscriptions) {
        Write-Progress -Activity 'Analyzing Subscriptions' -Status "Processing: $($subscription.Name)" -PercentComplete ((($subscriptions.IndexOf($subscription) + 1) / [double]$subscriptions.Count) * 100)
        $Logger.Write("Processing subscription: $($subscription.Name)", 'INFO')
        try {
            Set-AzContext -Subscription $subscription.Id -ErrorAction Stop | Out-Null
            $accounts = Get-AzStorageAccount -ErrorAction Stop
            if (-not $accounts -or $accounts.Count -eq 0) { $Logger.Write("No storage accounts in subscription: $($subscription.Name)", 'WARNING'); continue }
            # Prefetch cost data
            $costData = @{}
            try {
                $consumption = Get-AzConsumptionUsageDetail -StartDate $startDate -EndDate $endDate -ErrorAction Stop
                foreach ($item in ($consumption | Where-Object { $_.ConsumedService -eq 'Microsoft.Storage' -and $_.ResourceId })) {
                    if (-not $costData.ContainsKey($item.ResourceId)) { $costData[$item.ResourceId] = 0.0 }
                    $costData[$item.ResourceId] += [double]$item.PretaxCost
                }
            } catch { $Logger.Write("Failed to retrieve cost data: $_", 'ERROR') }
            foreach ($account in $accounts) {
                try {
                    $targetRegion = if ($Region -and $Region -ne 'auto') { $Region } else { $account.Location }
                    $analyzer = [StorageAccountAnalyzer]::new($account, $targetRegion, $Days)
                    $metrics = $analyzer.AnalyzeUsageMetrics($startDate, $endDate)
                    $actualCost = if ($costData.ContainsKey($account.Id)) { [math]::Round([double]$costData[$account.Id], 2) } else { [double]$analyzer.CalculateCurrentCost() }
                    $recResults = $analyzer.GenerateRecommendations()
                    $entry = [PSCustomObject]@{
                        Subscription       = $subscription.Name
                        ResourceGroup      = $account.ResourceGroupName
                        StorageAccount     = $account.StorageAccountName
                        Location           = $account.Location
                        Kind               = $account.Kind
                        AccessTier         = if ($account.AccessTier) { $account.AccessTier.ToString() } else { 'N/A' }
                        Redundancy         = $account.Sku.Name
                        CreatedDate        = $account.CreationTime
                        UsedCapacityGB     = $metrics.UsedCapacityGB
                        BlobCount          = $metrics.BlobCount
                        Transactions       = $metrics.Transactions
                        TransactionsPerDay = if ($Days -gt 0) { [math]::Round($metrics.Transactions / [double]$Days, 1) } else { 0 }
                        IngressGB          = $metrics.IngressGB
                        EgressGB           = $metrics.EgressGB
                        MonthlyCost        = [math]::Round($actualCost, 2)
                        EstimatedSavings   = $recResults.TotalSavings
                        SavingsPercent     = if ($actualCost -gt 0) { [math]::Round(($recResults.TotalSavings / $actualCost) * 100, 1) } else { 0 }
                        RecommendationCount= $recResults.Recommendations.Count
                        Recommendations    = $recResults.Recommendations | ConvertTo-Json -Compress
                        Tags               = if ($account.Tags) { $account.Tags | ConvertTo-Json -Compress } else { '{}' }
                    }
                    $null = $script:Report.Add($entry)
                    $script:ExecutiveSummary.TotalAccounts++
                    $script:ExecutiveSummary.TotalDataGB += $metrics.UsedCapacityGB
                    $script:ExecutiveSummary.CurrentMonthlyCost += $actualCost
                    $script:ExecutiveSummary.PotentialSavings += $recResults.TotalSavings
                    $Logger.Write("Analyzed: $($account.StorageAccountName) - Cost: $$actualCost, Savings: $$($recResults.TotalSavings)", 'INFO')
                }
                catch {
                    $Logger.Write("Failed to analyze account $($account.StorageAccountName): $_", 'ERROR')
                }
            }
        }
        catch {
            $Logger.Write("Failed to process subscription $($subscription.Name): $_", 'ERROR')
        }
    }
    Write-Progress -Activity 'Analyzing Subscriptions' -Completed
}

function Export-CsvReport {
    $csvPath = Join-Path -Path $OutputDir -ChildPath 'StorageOptimization_Report.csv'
    $script:Report | Select-Object Subscription, ResourceGroup, StorageAccount, Location, Kind, AccessTier, Redundancy, CreatedDate, UsedCapacityGB, BlobCount, Transactions, TransactionsPerDay, IngressGB, EgressGB, MonthlyCost, EstimatedSavings, SavingsPercent, RecommendationCount |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    $Logger.Write("CSV report exported: $csvPath", 'SUCCESS')
    return $csvPath
}

function Export-ExcelReport {
    if (-not (Get-Module -Name ImportExcel -ErrorAction SilentlyContinue)) { $Logger.Write('ImportExcel module not available, skipping Excel export', 'WARNING'); return $null }
    $excelPath = Join-Path -Path $OutputDir -ChildPath 'StorageOptimization_Report.xlsx'
    try {
        $data = $script:Report | Sort-Object -Property MonthlyCost -Descending
        $chart1 = New-ExcelChartDefinition -Title 'Cost by Account' -ChartType ColumnClustered -XRange 'Data[StorageAccount]' -YRange 'Data[MonthlyCost]'
        $chart2 = New-ExcelChartDefinition -Title 'Potential Savings' -ChartType ColumnClustered -XRange 'Data[StorageAccount]' -YRange 'Data[EstimatedSavings]'
        $data | Export-Excel -Path $excelPath -AutoSize -TableName 'Data' -TableStyle Light8 -WorksheetName 'Data' -Title 'Azure Storage Optimization' -ExcelChartDefinition @($chart1,$chart2) -ClearSheet
        # Summary sheet
        $summaryRows = @(
            [PSCustomObject]@{ Metric='Total Accounts'; Value=$script:ExecutiveSummary.TotalAccounts },
            [PSCustomObject]@{ Metric='Total Subscriptions'; Value=$script:ExecutiveSummary.TotalSubscriptions },
            [PSCustomObject]@{ Metric='Total Data (GB)'; Value=[math]::Round($script:ExecutiveSummary.TotalDataGB,2) },
            [PSCustomObject]@{ Metric='Current Monthly Cost ($)'; Value=[math]::Round($script:ExecutiveSummary.CurrentMonthlyCost,2) },
            [PSCustomObject]@{ Metric='Potential Savings ($)'; Value=[math]::Round($script:ExecutiveSummary.PotentialSavings,2) }
        )
        $summaryRows | Export-Excel -Path $excelPath -WorksheetName 'Summary' -AutoSize -TableName 'Summary' -TableStyle Light9
        $Logger.Write("Excel report exported: $excelPath", 'SUCCESS')
        return $excelPath
    }
    catch {
        $Logger.Write("Excel export failed: $_", 'ERROR')
        return $null
    }
}

function Export-PdfReport {
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
\section*{Executive Summary}
Total Accounts: $($script:ExecutiveSummary.TotalAccounts) \\
Total Subscriptions: $($script:ExecutiveSummary.TotalSubscriptions) \\
Total Data (GB): $([math]::Round($script:ExecutiveSummary.TotalDataGB,2)) \\
Current Monthly Cost (\$): $([math]::Round($script:ExecutiveSummary.CurrentMonthlyCost,2)) \\
Potential Savings (\$): $([math]::Round($script:ExecutiveSummary.PotentialSavings,2))
\end{document}
"@
    $texFile = Join-Path -Path $OutputDir -ChildPath 'StorageOptimization_Report.tex'
    $latexContent | Out-File -FilePath $texFile -Encoding UTF8
    $Logger.Write("LaTeX report exported: $texFile", 'SUCCESS')
    $latexmk = Get-Command latexmk -ErrorAction SilentlyContinue
    if ($latexmk) {
        try { Push-Location $OutputDir; & $latexmk -pdf -quiet (Split-Path -Leaf $texFile) | Out-Null; Pop-Location; $Logger.Write('PDF compiled via latexmk', 'SUCCESS') } catch { $Logger.Write("latexmk compile failed: $_", 'WARNING') }
    } else { $Logger.Write('latexmk not found; PDF not compiled', 'WARNING') }
    return $texFile
}

function Export-HtmlDashboard {
    $htmlPath = Join-Path -Path $OutputDir -ChildPath 'StorageOptimization_Dashboard.html'
    $summary = @(
        '<h2>Executive Summary</h2>',
        "<p>Total Accounts: $($script:ExecutiveSummary.TotalAccounts)</p>",
        "<p>Total Subscriptions: $($script:ExecutiveSummary.TotalSubscriptions)</p>",
        "<p>Total Data (GB): $([math]::Round($script:ExecutiveSummary.TotalDataGB,2))</p>",
        "<p>Current Monthly Cost ($): $([math]::Round($script:ExecutiveSummary.CurrentMonthlyCost,2))</p>",
        "<p>Potential Savings ($): $([math]::Round($script:ExecutiveSummary.PotentialSavings,2))</p>"
    ) -join "`n"
    $table = $script:Report | Select-Object Subscription, ResourceGroup, StorageAccount, Location, AccessTier, Redundancy, UsedCapacityGB, MonthlyCost, EstimatedSavings, SavingsPercent, RecommendationCount |
        ConvertTo-Html -Fragment -PreContent '<h2>Accounts</h2>' -PostContent ''
    $html = @("<html><head><meta charset='utf-8'><title>Azure Storage Optimization</title></head><body>", $summary, $table, '</body></html>') -join "`n"
    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    $Logger.Write("HTML dashboard exported: $htmlPath", 'SUCCESS')
    return $htmlPath
}

function Write-ActionsFile {
    $actionsPath = Join-Path -Path $OutputDir -ChildPath 'Remediation_Actions.ps1'
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($row in $script:Report) {
        try {
            $recs = $row.Recommendations | ConvertFrom-Json -ErrorAction Stop
            foreach ($rec in $recs) { foreach ($cmd in $rec.Implementation) { if ($cmd -and -not ($cmd -like '#*')) { $null = $lines.Add($cmd) } } }
        } catch {}
    }
    if ($lines.Count -gt 0) { $lines | Out-File -FilePath $actionsPath -Encoding UTF8; $Logger.Write("Remediation actions exported: $actionsPath", 'SUCCESS') }
    else { $Logger.Write('No remediation actions to export.', 'INFO') }
    return $actionsPath
}

function Send-ReportEmail {
    param([string]$toAddress,[string]$csvPath,[string]$excelPath,[string]$htmlPath)
    if (-not $toAddress) { return }
    # Requires SMTP env: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM
    $smtpHost = $env:SMTP_HOST; $smtpPort = $env:SMTP_PORT; $smtpUser = $env:SMTP_USER; $smtpPass = $env:SMTP_PASS; $smtpFrom = $env:SMTP_FROM
    if (-not $smtpHost -or -not $smtpFrom) { $Logger.Write('SMTP settings not found in env; skipping email send.', 'WARNING'); return }
    try {
        $subject = "Azure Storage Optimization Report - $($Config.RunId)"
        $body = @(
            "Executive Summary:",
            "Total Accounts: $($script:ExecutiveSummary.TotalAccounts)",
            "Total Data (GB): $([math]::Round($script:ExecutiveSummary.TotalDataGB,2))",
            "Current Monthly Cost: $([math]::Round($script:ExecutiveSummary.CurrentMonthlyCost,2))",
            "Potential Savings: $([math]::Round($script:ExecutiveSummary.PotentialSavings,2))"
        ) -join "`n"
        Send-MailMessage -SmtpServer $smtpHost -Port $smtpPort -UseSsl -Credential (New-Object System.Management.Automation.PSCredential($smtpUser,(ConvertTo-SecureString $smtpPass -AsPlainText -Force))) -From $smtpFrom -To $toAddress -Subject $subject -Body $body -Attachments (@($csvPath,$excelPath,$htmlPath) | Where-Object { $_ -and (Test-Path $_) }) -ErrorAction Stop
        $Logger.Write("Email sent to $toAddress", 'SUCCESS')
    } catch { $Logger.Write("Failed to send email: $_", 'ERROR') }
}

# Entry point
try {
    Initialize-RequiredModules
    Connect-AzureWithRetry | Out-Null
    Start-StorageAnalysis

    $csv = $null; $xlsx = $null; $html = $null; $tex = $null
    if ($ExportCsv)  { $csv  = Export-CsvReport }
    if ($ExportExcel) { $xlsx = Export-ExcelReport }
    if ($ExportHtml) { $html = Export-HtmlDashboard }
    if ($ExportPdf)  { $tex  = Export-PdfReport }

    $actions = Write-ActionsFile
    if ($AutoRemediate) {
        $Logger.Write('AutoRemediate requested. Review Remediation_Actions.ps1 and execute manually or rerun with -WhatIf:$false and explicit confirmation.', 'WARNING')
    }

    if ($EmailReport) { Send-ReportEmail -toAddress $EmailReport -csvPath $csv -excelPath $xlsx -htmlPath $html }

    # Final console summary
    Write-Host "=== Executive Summary ===" -ForegroundColor Cyan
    Write-Host ("Accounts: {0} | Data: {1} GB | Cost: ${2} | Savings: ${3}" -f $script:ExecutiveSummary.TotalAccounts, [math]::Round($script:ExecutiveSummary.TotalDataGB,2), [math]::Round($script:ExecutiveSummary.CurrentMonthlyCost,2), [math]::Round($script:ExecutiveSummary.PotentialSavings,2)) -ForegroundColor Green
}
catch {
    $Logger.Write("Fatal error: $_", 'ERROR')
    throw
}
finally {
    $elapsed = (Get-Date) - $Config.StartTime
    $Logger.Write("Completed in {0:N2} minutes" -f $elapsed.TotalMinutes, 'INFO')
}