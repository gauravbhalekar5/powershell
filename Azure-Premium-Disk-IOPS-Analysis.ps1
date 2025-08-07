# Azure Premium Disk IOPS Analysis Script
# Analyzes IOPS usage across all subscriptions for premium disks
# Generates recommendations for premium to standard disk migration
# Compatible with PowerShell ISE and Console

param(
    [string[]]$TargetSubscriptions = @(  # Specific DENV subscriptions to analyze
        "DENV manufacturing prod",
        "DENV Prod",
        "Daikin Europe",
        "DENV data-analytics PROD",
        "DENV hub",
        "DENV non-prod",
        "DENV prod",
        "DENV data-analytics DEV"
    ),
    [int]$AnalysisDays = 7,              # Days to analyze
    [double]$UtilizationThreshold = 30,  # IOPS utilization threshold (%)
    [string]$OutputPath = ".\DENV_DiskAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [switch]$ExportToExcel = $true,      # Export to Excel format (default enabled)
    [switch]$IncludeCurrentMetrics,      # Include real-time metrics
    [string]$LogAnalyticsWorkspace = "", # Log Analytics Workspace ID for historical data
    [switch]$DetailedReport = $true      # Generate detailed cost analysis report (default enabled)
)

# Check if running in PowerShell ISE
$isISE = $psISE -ne $null

Write-Host "Azure Premium Disk IOPS Analysis Tool - DENV Environment" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "Target: DENV and Daikin Europe subscriptions only" -ForegroundColor Yellow
Write-Host "Purpose: Identify underutilized premium disks for cost optimization" -ForegroundColor Yellow
if ($isISE) { Write-Host "Running in PowerShell ISE" -ForegroundColor Green }
Write-Host ""

# Premium disk IOPS and pricing information (Updated with correct specifications)
$PremiumDiskSpecs = @{
    'P1'   = @{ SizeGB = 4;     BaseIOPS = 120;   BurstIOPS = 3500;   Throughput = 25;   MonthlyCost = 0.6 }
    'P2'   = @{ SizeGB = 8;     BaseIOPS = 120;   BurstIOPS = 3500;   Throughput = 25;   MonthlyCost = 1.2 }
    'P3'   = @{ SizeGB = 16;    BaseIOPS = 120;   BurstIOPS = 3500;   Throughput = 25;   MonthlyCost = 2.3 }
    'P4'   = @{ SizeGB = 32;    BaseIOPS = 120;   BurstIOPS = 3500;   Throughput = 25;   MonthlyCost = 4.6 }
    'P6'   = @{ SizeGB = 64;    BaseIOPS = 240;   BurstIOPS = 3500;   Throughput = 50;   MonthlyCost = 9.2 }
    'P10'  = @{ SizeGB = 128;   BaseIOPS = 500;   BurstIOPS = 3500;   Throughput = 100;  MonthlyCost = 18.4 }
    'P15'  = @{ SizeGB = 256;   BaseIOPS = 1100;  BurstIOPS = 3500;   Throughput = 125;  MonthlyCost = 36.8 }
    'P20'  = @{ SizeGB = 512;   BaseIOPS = 2300;  BurstIOPS = 3500;   Throughput = 150;  MonthlyCost = 73.6 }
    'P30'  = @{ SizeGB = 1024;  BaseIOPS = 5000;  BurstIOPS = 5000;   Throughput = 200;  MonthlyCost = 147.2 }
    'P40'  = @{ SizeGB = 2048;  BaseIOPS = 7500;  BurstIOPS = 7500;   Throughput = 250;  MonthlyCost = 294.4 }
    'P50'  = @{ SizeGB = 4096;  BaseIOPS = 7500;  BurstIOPS = 7500;   Throughput = 250;  MonthlyCost = 588.8 }
    'P60'  = @{ SizeGB = 8192;  BaseIOPS = 16000; BurstIOPS = 16000;  Throughput = 400;  MonthlyCost = 1177.6 }
    'P70'  = @{ SizeGB = 16384; BaseIOPS = 18000; BurstIOPS = 18000;  Throughput = 500;  MonthlyCost = 2355.2 }
    'P80'  = @{ SizeGB = 32768; BaseIOPS = 20000; BurstIOPS = 20000;  Throughput = 750;  MonthlyCost = 4710.4 }
}

$StandardDiskSpecs = @{
    'S4'   = @{ SizeGB = 32;    IOPS = 500;   Throughput = 60;   MonthlyCost = 1.54 }
    'S6'   = @{ SizeGB = 64;    IOPS = 500;   Throughput = 60;   MonthlyCost = 3.01 }
    'S10'  = @{ SizeGB = 128;   IOPS = 500;   Throughput = 60;   MonthlyCost = 5.89 }
    'S15'  = @{ SizeGB = 256;   IOPS = 500;   Throughput = 60;   MonthlyCost = 11.52 }
    'S20'  = @{ SizeGB = 512;   IOPS = 500;   Throughput = 60;   MonthlyCost = 22.56 }
    'S30'  = @{ SizeGB = 1024;  IOPS = 500;   Throughput = 60;   MonthlyCost = 44.16 }
    'S40'  = @{ SizeGB = 2048;  IOPS = 500;   Throughput = 60;   MonthlyCost = 86.40 }
    'S50'  = @{ SizeGB = 4096;  IOPS = 500;   Throughput = 60;   MonthlyCost = 168.96 }
    'S60'  = @{ SizeGB = 8192;  IOPS = 1300;  Throughput = 300;  MonthlyCost = 330.24 }
    'S70'  = @{ SizeGB = 16384; IOPS = 2000;  Throughput = 500;  MonthlyCost = 645.12 }
    'S80'  = @{ SizeGB = 32768; IOPS = 2000;  Throughput = 500;  MonthlyCost = 1260.48 }
}

# Function to check Azure PowerShell module
function Test-AzureModule {
    try {
        # Check for required modules
        $requiredModules = @('Az.Accounts', 'Az.Compute', 'Az.Storage', 'Az.Resources', 'Az.Monitor')
        $missingModules = @()
        
        foreach ($module in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                $missingModules += $module
            }
        }
        
        if ($missingModules.Count -gt 0) {
            Write-Warning "Missing required modules: $($missingModules -join ', ')"
            Write-Host "Install missing modules with:" -ForegroundColor Yellow
            Write-Host "Install-Module -Name $($missingModules -join ', ') -Force -AllowClobber" -ForegroundColor Yellow
            return $false
        }
        
        # Import modules
        foreach ($module in $requiredModules) {
            Import-Module $module -Force -ErrorAction SilentlyContinue
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to verify Azure PowerShell modules: $($_.Exception.Message)"
        return $false
    }
}

# Function to connect to Azure
function Connect-ToAzure {
    try {
        $context = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $context -or -not $context.Account) {
            Write-Host "Connecting to Azure..." -ForegroundColor Yellow
            $connection = Connect-AzAccount -ErrorAction Stop
            if (-not $connection) {
                throw "Failed to establish Azure connection"
            }
        } else {
            Write-Host "Already connected to Azure as: $($context.Account.Id)" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Error "Failed to connect to Azure: $($_.Exception.Message)"
        return $false
    }
}

# Function to get specific DENV subscriptions
function Get-TargetSubscriptions {
    param([string[]]$TargetNames)
    
    try {
        Write-Host "Searching for target subscriptions..." -ForegroundColor Yellow
        $allSubscriptions = Get-AzSubscription -ErrorAction Stop
        $foundSubscriptions = @()
        $notFoundSubscriptions = @()
        
        foreach ($targetName in $TargetNames) {
            $matchedSub = $allSubscriptions | Where-Object {
                $_.Name -eq $targetName -or
                $_.Name -like "*$targetName*" -or
                $targetName -like "*$($_.Name)*"
            }
            
            if ($matchedSub) {
                if ($matchedSub -is [array]) {
                    # Multiple matches, take the exact match first
                    $exactMatch = $matchedSub | Where-Object { $_.Name -eq $targetName }
                    if ($exactMatch) {
                        $foundSubscriptions += $exactMatch
                        Write-Host "  ✓ Found exact match: $($exactMatch.Name) ($($exactMatch.Id))" -ForegroundColor Green
                    } else {
                        $foundSubscriptions += $matchedSub[0]
                        Write-Host "  ✓ Found similar: $($matchedSub[0].Name) for target '$targetName'" -ForegroundColor Yellow
                    }
                } else {
                    $foundSubscriptions += $matchedSub
                    Write-Host "  ✓ Found: $($matchedSub.Name) ($($matchedSub.Id))" -ForegroundColor Green
                }
            } else {
                $notFoundSubscriptions += $targetName
                Write-Host "  ✗ Not found: $targetName" -ForegroundColor Red
            }
        }
        
        if ($notFoundSubscriptions.Count -gt 0) {
            Write-Host ""
            Write-Host "Available subscriptions for reference:" -ForegroundColor Gray
            $allSubscriptions | Select-Object -First 10 | ForEach-Object { 
                Write-Host "  - $($_.Name)" -ForegroundColor Gray 
            }
            if ($allSubscriptions.Count -gt 10) {
                Write-Host "  ... and $($allSubscriptions.Count - 10) more" -ForegroundColor Gray
            }
        }
        
        Write-Host ""
        Write-Host "Final analysis will include $($foundSubscriptions.Count) subscription(s):" -ForegroundColor Cyan
        $foundSubscriptions | ForEach-Object { Write-Host "  → $($_.Name)" -ForegroundColor White }
        
        return $foundSubscriptions
    }
    catch {
        Write-Error "Failed to get target subscriptions: $($_.Exception.Message)"
        return @()
    }
}

# Function to determine disk tier based on size
function Get-DiskTier {
    param([int]$DiskSizeGB, [string]$SkuName)
    
    if ($SkuName -like "Premium*") {
        # Premium disk tier determination
        $sortedTiers = $PremiumDiskSpecs.Keys | Sort-Object { $PremiumDiskSpecs[$_].SizeGB }
        foreach ($tier in $sortedTiers) {
            if ($DiskSizeGB -le $PremiumDiskSpecs[$tier].SizeGB) {
                return $tier
            }
        }
        return "P80" # Fallback to largest tier
    } else {
        # Standard disk tier determination
        $sortedTiers = $StandardDiskSpecs.Keys | Sort-Object { $StandardDiskSpecs[$_].SizeGB }
        foreach ($tier in $sortedTiers) {
            if ($DiskSizeGB -le $StandardDiskSpecs[$tier].SizeGB) {
                return $tier
            }
        }
        return "S80" # Fallback to largest tier
    }
}

# Function to get premium disks from a subscription
function Get-PremiumDisksFromSubscription {
    param([string]$SubscriptionId, [string]$SubscriptionName)
    
    try {
        Write-Host "Analyzing subscription: $SubscriptionName" -ForegroundColor Cyan
        $null = Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop
        
        # Get all managed disks
        Write-Host "  Retrieving disk information..." -ForegroundColor Gray
        $allDisks = Get-AzDisk -ErrorAction Stop
        $premiumDisks = $allDisks | Where-Object { $_.Sku.Name -like "Premium*" }
        
        Write-Host "  Found $($premiumDisks.Count) premium disk(s)" -ForegroundColor Gray
        
        $diskDetails = @()
        foreach ($disk in $premiumDisks) {
            try {
                # Get VM information if disk is attached
                $vmInfo = $null
                $vmName = "Not Attached"
                $vmResourceGroup = "N/A"
                $vmPowerState = "N/A"
                
                if ($disk.ManagedBy) {
                    $vmResourceId = $disk.ManagedBy
                    $vmName = ($vmResourceId -split "/")[-1]
                    $vmResourceGroup = ($vmResourceId -split "/")[4]
                    
                    try {
                        $vmInfo = Get-AzVM -ResourceGroupName $vmResourceGroup -Name $vmName -Status -ErrorAction SilentlyContinue
                        if ($vmInfo) {
                            $vmPowerState = ($vmInfo.Statuses | Where-Object { $_.Code -like "PowerState/*" }).DisplayStatus
                        }
                    } catch {
                        Write-Warning "Could not get VM info for $vmName"
                    }
                }
                
                # Determine disk tier
                $diskTier = Get-DiskTier -DiskSizeGB $disk.DiskSizeGB -SkuName $disk.Sku.Name
                $provisionedBaseIOPS = if ($PremiumDiskSpecs.ContainsKey($diskTier)) { $PremiumDiskSpecs[$diskTier].BaseIOPS } else { 0 }
                $provisionedBurstIOPS = if ($PremiumDiskSpecs.ContainsKey($diskTier)) { $PremiumDiskSpecs[$diskTier].BurstIOPS } else { 0 }
                $provisionedThroughput = if ($PremiumDiskSpecs.ContainsKey($diskTier)) { $PremiumDiskSpecs[$diskTier].Throughput } else { 0 }
                
                $diskDetail = [PSCustomObject]@{
                    SubscriptionName = $SubscriptionName
                    SubscriptionId = $SubscriptionId
                    ResourceGroup = $disk.ResourceGroupName
                    DiskName = $disk.Name
                    DiskId = $disk.Id
                    DiskTier = $diskTier
                    DiskSizeGB = $disk.DiskSizeGB
                    SkuName = $disk.Sku.Name
                    ProvisionedBaseIOPS = $provisionedBaseIOPS
                    ProvisionedBurstIOPS = $provisionedBurstIOPS
                    ProvisionedThroughputMBps = $provisionedThroughput
                    AttachedVM = $vmName
                    VMResourceGroup = $vmResourceGroup
                    VMPowerState = $vmPowerState
                    DiskState = $disk.DiskState
                    Location = $disk.Location
                    CreatedTime = $disk.TimeCreated
                    EstimatedMonthlyCost = if ($PremiumDiskSpecs.ContainsKey($diskTier)) { $PremiumDiskSpecs[$diskTier].MonthlyCost } else { 0 }
                    Tags = if ($disk.Tags) { ($disk.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " } else { "" }
                }
                
                $diskDetails += $diskDetail
            }
            catch {
                Write-Warning "Error processing disk $($disk.Name): $($_.Exception.Message)"
            }
        }
        
        return $diskDetails
    }
    catch {
        Write-Warning "Error analyzing subscription $SubscriptionName : $($_.Exception.Message)"
        return @()
    }
}

# Function to get real disk metrics using Azure Monitor
function Get-RealDiskMetrics {
    param(
        [object]$DiskInfo, 
        [int]$Days,
        [string]$SubscriptionId
    )
    
    try {
        $endTime = Get-Date
        $startTime = $endTime.AddDays(-$Days)
        
        # Set context for metrics query
        $null = Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction SilentlyContinue
        
        $metrics = @{
            AvgIOPS = 0
            MaxIOPS = 0
            AvgReadIOPS = 0
            AvgWriteIOPS = 0
            AvgThroughputMBps = 0
            MaxThroughputMBps = 0
            AvgLatencyMs = 0
            DataPointsCollected = 0
        }
        
        # Only get metrics if disk is attached to a running VM
        if ($DiskInfo.AttachedVM -ne "Not Attached" -and $DiskInfo.VMPowerState -like "*running*") {
            try {
                # Get IOPS metrics
                $iopsReadMetric = Get-AzMetric -ResourceId $DiskInfo.DiskId -MetricName "Disk Read Operations/Sec" -StartTime $startTime -EndTime $endTime -TimeGrain "01:00:00" -ErrorAction SilentlyContinue
                $iopsWriteMetric = Get-AzMetric -ResourceId $DiskInfo.DiskId -MetricName "Disk Write Operations/Sec" -StartTime $startTime -EndTime $endTime -TimeGrain "01:00:00" -ErrorAction SilentlyContinue
                
                if ($iopsReadMetric -and $iopsReadMetric.Data) {
                    $readIOPSValues = $iopsReadMetric.Data | Where-Object { $_.Average -ne $null } | ForEach-Object { $_.Average }
                    if ($readIOPSValues) {
                        $metrics.AvgReadIOPS = [math]::Round(($readIOPSValues | Measure-Object -Average).Average, 2)
                        $metrics.DataPointsCollected += $readIOPSValues.Count
                    }
                }
                
                if ($iopsWriteMetric -and $iopsWriteMetric.Data) {
                    $writeIOPSValues = $iopsWriteMetric.Data | Where-Object { $_.Average -ne $null } | ForEach-Object { $_.Average }
                    if ($writeIOPSValues) {
                        $metrics.AvgWriteIOPS = [math]::Round(($writeIOPSValues | Measure-Object -Average).Average, 2)
                        $metrics.DataPointsCollected += $writeIOPSValues.Count
                    }
                }
                
                $metrics.AvgIOPS = $metrics.AvgReadIOPS + $metrics.AvgWriteIOPS
                $metrics.MaxIOPS = [math]::Max($metrics.AvgReadIOPS, $metrics.AvgWriteIOPS) * 2 # Rough estimate
                
                # Get throughput metrics
                $readBytesMetric = Get-AzMetric -ResourceId $DiskInfo.DiskId -MetricName "Disk Read Bytes/sec" -StartTime $startTime -EndTime $endTime -TimeGrain "01:00:00" -ErrorAction SilentlyContinue
                $writeBytesMetric = Get-AzMetric -ResourceId $DiskInfo.DiskId -MetricName "Disk Write Bytes/sec" -StartTime $startTime -EndTime $endTime -TimeGrain "01:00:00" -ErrorAction SilentlyContinue
                
                if ($readBytesMetric -and $readBytesMetric.Data) {
                    $readBytesValues = $readBytesMetric.Data | Where-Object { $_.Average -ne $null } | ForEach-Object { $_.Average / 1024 / 1024 } # Convert to MB/s
                    if ($readBytesValues) {
                        $avgReadMBps = ($readBytesValues | Measure-Object -Average).Average
                        $metrics.AvgThroughputMBps += $avgReadMBps
                        $metrics.MaxThroughputMBps = [math]::Max($metrics.MaxThroughputMBps, ($readBytesValues | Measure-Object -Maximum).Maximum)
                    }
                }
                
                if ($writeBytesMetric -and $writeBytesMetric.Data) {
                    $writeBytesValues = $writeBytesMetric.Data | Where-Object { $_.Average -ne $null } | ForEach-Object { $_.Average / 1024 / 1024 } # Convert to MB/s
                    if ($writeBytesValues) {
                        $avgWriteMBps = ($writeBytesValues | Measure-Object -Average).Average
                        $metrics.AvgThroughputMBps += $avgWriteMBps
                        $metrics.MaxThroughputMBps = [math]::Max($metrics.MaxThroughputMBps, ($writeBytesValues | Measure-Object -Maximum).Maximum)
                    }
                }
                
                $metrics.AvgThroughputMBps = [math]::Round($metrics.AvgThroughputMBps, 2)
                $metrics.MaxThroughputMBps = [math]::Round($metrics.MaxThroughputMBps, 2)
                
            } catch {
                Write-Warning "Could not retrieve metrics for disk $($DiskInfo.DiskName): $($_.Exception.Message)"
            }
        }
        
        # If no real metrics available, use conservative estimates based on disk state
        if ($metrics.DataPointsCollected -eq 0) {
            if ($DiskInfo.AttachedVM -eq "Not Attached") {
                # Unattached disk - zero usage
                $metrics.AvgIOPS = 0
                $metrics.MaxIOPS = 0
                $metrics.AvgThroughputMBps = 0
            } elseif ($DiskInfo.VMPowerState -notlike "*running*") {
                # VM not running - minimal usage
                $metrics.AvgIOPS = [math]::Round($DiskInfo.ProvisionedBaseIOPS * 0.05, 0) # 5% of provisioned
                $metrics.MaxIOPS = [math]::Round($DiskInfo.ProvisionedBaseIOPS * 0.15, 0) # 15% of provisioned
                $metrics.AvgThroughputMBps = [math]::Round($DiskInfo.ProvisionedThroughputMBps * 0.05, 2)
            } else {
                # VM running but no metrics - conservative estimate
                $metrics.AvgIOPS = [math]::Round($DiskInfo.ProvisionedBaseIOPS * 0.20, 0) # 20% of provisioned
                $metrics.MaxIOPS = [math]::Round($DiskInfo.ProvisionedBaseIOPS * 0.40, 0) # 40% of provisioned
                $metrics.AvgThroughputMBps = [math]::Round($DiskInfo.ProvisionedThroughputMBps * 0.20, 2)
            }
        }
        
        return $metrics
    }
    catch {
        Write-Warning "Error getting metrics for disk $($DiskInfo.DiskName): $($_.Exception.Message)"
        # Return zero metrics on error
        return @{
            AvgIOPS = 0; MaxIOPS = 0; AvgReadIOPS = 0; AvgWriteIOPS = 0
            AvgThroughputMBps = 0; MaxThroughputMBps = 0; AvgLatencyMs = 0
            DataPointsCollected = 0
        }
    }
}

# Function to generate migration recommendations
function Get-MigrationRecommendation {
    param([object]$DiskInfo, [object]$Metrics, [double]$Threshold)
    
    $baseUtilization = if ($DiskInfo.ProvisionedBaseIOPS -gt 0) {
        ($Metrics.AvgIOPS / $DiskInfo.ProvisionedBaseIOPS) * 100
    } else { 0 }
    
    $burstUtilization = if ($DiskInfo.ProvisionedBurstIOPS -gt 0) {
        ($Metrics.MaxIOPS / $DiskInfo.ProvisionedBurstIOPS) * 100
    } else { 0 }
    
    $recommendation = "Keep Premium"
    $recommendedTier = $DiskInfo.DiskTier
    $potentialSavings = 0
    $recommendationReason = "High utilization"
    
    # Decision logic for migration
    if ($DiskInfo.AttachedVM -eq "Not Attached") {
        $recommendation = "Consider Deletion or Archive"
        $recommendationReason = "Disk not attached to any VM"
        $potentialSavings = $DiskInfo.EstimatedMonthlyCost
    }
    elseif ($DiskInfo.VMPowerState -notlike "*running*" -and $DiskInfo.VMPowerState -ne "N/A") {
        $recommendation = "Migrate to Standard"
        $recommendationReason = "VM not running - minimal performance requirements"
        # Find suitable standard disk
        $suitableStandardTier = $null
        foreach ($tier in ($StandardDiskSpecs.Keys | Sort-Object { $StandardDiskSpecs[$_].SizeGB })) {
            $spec = $StandardDiskSpecs[$tier]
            if ($spec.SizeGB -ge $DiskInfo.DiskSizeGB) {
                $suitableStandardTier = $tier
                break
            }
        }
        if ($suitableStandardTier) {
            $recommendedTier = $suitableStandardTier
            $potentialSavings = $DiskInfo.EstimatedMonthlyCost - $StandardDiskSpecs[$suitableStandardTier].MonthlyCost
        }
    }
    elseif ($baseUtilization -lt $Threshold) {
        # Find appropriate standard disk that can handle the workload
        $suitableStandardTier = $null
        foreach ($tier in ($StandardDiskSpecs.Keys | Sort-Object { $StandardDiskSpecs[$_].SizeGB })) {
            $spec = $StandardDiskSpecs[$tier]
            if ($spec.SizeGB -ge $DiskInfo.DiskSizeGB -and 
                $spec.IOPS -ge $Metrics.MaxIOPS -and 
                $spec.Throughput -ge $Metrics.MaxThroughputMBps) {
                $suitableStandardTier = $tier
                break
            }
        }
        
        if ($suitableStandardTier) {
            $recommendation = "Migrate to Standard"
            $recommendedTier = $suitableStandardTier
            $potentialSavings = $DiskInfo.EstimatedMonthlyCost - $StandardDiskSpecs[$suitableStandardTier].MonthlyCost
            $recommendationReason = "Low utilization ($([math]::Round($baseUtilization, 1))% of base IOPS)"
        } else {
            $recommendation = "Keep Premium"
            $recommendationReason = "No suitable standard disk tier available for workload requirements"
        }
    } else {
        $recommendationReason = "Utilization ($([math]::Round($baseUtilization, 1))%) above threshold"
    }
    
    return @{
        Recommendation = $recommendation
        RecommendedTier = $recommendedTier
        BaseUtilizationPercent = [math]::Round($baseUtilization, 2)
        BurstUtilizationPercent = [math]::Round($burstUtilization, 2)
        PotentialMonthlySavings = [math]::Round([math]::Max($potentialSavings, 0), 2)
        RecommendationReason = $recommendationReason
    }
}

# Function to export results with improved error handling
function Export-Results {
    param([array]$Results, [string]$Path, [switch]$ToExcel)
    
    try {
        # Ensure output directory exists
        $outputDir = Split-Path $Path -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        if ($ToExcel) {
            try {
                # Try to use ImportExcel module
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    Import-Module ImportExcel -Force
                    $excelPath = "$Path.xlsx"
                    
                    # Create summary data
                    $summary = $Results | Group-Object Recommendation | Select-Object @{
                        Name = "Recommendation"
                        Expression = { $_.Name }
                    }, @{
                        Name = "Count"
                        Expression = { $_.Count }
                    }, @{
                        Name = "TotalMonthlyCost"
                        Expression = { [math]::Round(($_.Group | Measure-Object EstimatedMonthlyCost -Sum).Sum, 2) }
                    }, @{
                        Name = "PotentialSavings"
                        Expression = { [math]::Round(($_.Group | Measure-Object PotentialMonthlySavings -Sum).Sum, 2) }
                    }
                    
                    $subscriptionSummary = $Results | Group-Object SubscriptionName | Select-Object @{
                        Name = "SubscriptionName"
                        Expression = { $_.Name }
                    }, @{
                        Name = "TotalDisks"
                        Expression = { $_.Count }
                    }, @{
                        Name = "MigrationCandidates"
                        Expression = { ($_.Group | Where-Object {$_.Recommendation -eq 'Migrate to Standard'}).Count }
                    }, @{
                        Name = "DeletionCandidates"
                        Expression = { ($_.Group | Where-Object {$_.Recommendation -like '*Deletion*'}).Count }
                    }, @{
                        Name = "CurrentMonthlyCost"
                        Expression = { [math]::Round(($_.Group | Measure-Object EstimatedMonthlyCost -Sum).Sum, 2) }
                    }, @{
                        Name = "PotentialMonthlySavings"
                        Expression = { [math]::Round(($_.Group | Measure-Object PotentialMonthlySavings -Sum).Sum, 2) }
                    }
                    
                    # Export to Excel with multiple sheets
                    $Results | Export-Excel -Path $excelPath -WorksheetName "DiskAnalysis" -AutoSize -FreezeTopRow -BoldTopRow
                    $summary | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -FreezeTopRow -BoldTopRow
                    $subscriptionSummary | Export-Excel -Path $excelPath -WorksheetName "BySubscription" -AutoSize -FreezeTopRow -BoldTopRow
                    
                    Write-Host "Results exported to Excel: $excelPath" -ForegroundColor Green
                } else {
                    Write-Warning "ImportExcel module not available. Install with: Install-Module -Name ImportExcel"
                    throw "Excel export not available"
                }
            } catch {
                Write-Warning "Excel export failed: $($_.Exception.Message). Falling back to CSV."
                $ToExcel = $false
            }
        }
        
        if (-not $ToExcel) {
            # Export to CSV
            $csvPath = "$Path.csv"
            $Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Results exported to CSV: $csvPath" -ForegroundColor Green
        }
        
        # Generate comprehensive text report
        $reportPath = "$Path.txt"
        $totalDisks = $Results.Count
        $migrationCandidates = ($Results | Where-Object {$_.Recommendation -eq 'Migrate to Standard'}).Count
        $deletionCandidates = ($Results | Where-Object {$_.Recommendation -like '*Deletion*'}).Count
        $keepPremium = ($Results | Where-Object {$_.Recommendation -eq 'Keep Premium'}).Count
        $totalCurrentCost = ($Results | Measure-Object EstimatedMonthlyCost -Sum).Sum
        $totalPotentialSavings = ($Results | Measure-Object PotentialMonthlySavings -Sum).Sum
        $savingsPercentage = if ($totalCurrentCost -gt 0) { ($totalPotentialSavings / $totalCurrentCost) * 100 } else { 0 }
        
        $report = @"
DENV Azure Premium Disk Analysis Report
======================================
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Analysis Period: $AnalysisDays days
Utilization Threshold: $UtilizationThreshold%

TARGET SUBSCRIPTIONS:
$(($Results | Group-Object SubscriptionName | ForEach-Object { "- $($_.Name) ($($_.Count) disks)" }) -join "`n")

EXECUTIVE SUMMARY:
==================
Total Premium Disks Analyzed: $totalDisks
├─ Recommended for Migration to Standard: $migrationCandidates
├─ Consider Deletion/Archive: $deletionCandidates  
└─ Keep as Premium: $keepPremium

COST OPTIMIZATION OPPORTUNITY:
=============================
Current Monthly Premium Disk Cost: $([math]::Round($totalCurrentCost, 2)) USD
Potential Monthly Savings: $([math]::Round($totalPotentialSavings, 2)) USD
Potential Annual Savings: $([math]::Round($totalPotentialSavings * 12, 2)) USD
Savings Percentage: $([math]::Round($savingsPercentage, 1))%

SUBSCRIPTION BREAKDOWN:
======================
$($Results | Group-Object SubscriptionName | ForEach-Object {
    $subSavings = ($_.Group | Measure-Object PotentialMonthlySavings -Sum).Sum
    $subMigration = ($_.Group | Where-Object {$_.Recommendation -eq 'Migrate to Standard'}).Count
    $subDeletion = ($_.Group | Where-Object {$_.Recommendation -like '*Deletion*'}).Count
    "$($_.Name):"
    "  Total Disks: $($_.Count)"
    "  Migration Candidates: $subMigration"
    "  Deletion Candidates: $subDeletion"
    "  Monthly Savings: $([math]::Round($subSavings, 2)) USD"
    ""
})

TOP MIGRATION CANDIDATES:
========================
$($Results | Where-Object {$_.Recommendation -eq 'Migrate to Standard'} | Sort-Object PotentialMonthlySavings -Descending | Select-Object -First 10 | ForEach-Object {
    "• $($_.DiskName) ($($_.SubscriptionName))"
    "  Current: $($_.DiskTier) → Recommended: $($_.RecommendedTier)"
    "  Utilization: $($_.BaseUtilizationPercent)% | Savings: $($_.PotentialMonthlySavings) USD/month"
    "  Reason: $($_.RecommendationReason)"
    ""
})

UNATTACHED DISKS (Deletion Candidates):
======================================
$($Results | Where-Object {$_.Recommendation -like '*Deletion*'} | ForEach-Object {
    "• $($_.DiskName) ($($_.SubscriptionName))"
    "  Size: $($_.DiskSizeGB) GB | Cost: $($_.EstimatedMonthlyCost) USD/month"
    "  Created: $($_.CreatedTime)"
    ""
})

IMPLEMENTATION RECOMMENDATIONS:
==============================
1. IMMEDIATE ACTIONS:
   - Review unattached disks for deletion/archival
   - Validate migration candidates with application owners

2. MIGRATION STRATEGY:
   - Start with non-production environments
   - Migrate during maintenance windows
   - Monitor performance post-migration

3. ONGOING OPTIMIZATION:
   - Implement regular disk utilization monitoring
   - Set up alerts for underutilized premium disks
   - Review provisioning standards

4. BUSINESS JUSTIFICATION:
   - Low-risk optimization opportunity
   - No application downtime required
   - Immediate cost reduction
   - Supports cloud cost governance

TECHNICAL NOTES:
===============
- Metrics based on $AnalysisDays-day average usage
- Base IOPS utilization threshold: $UtilizationThreshold%
- Recommendations consider both IOPS and throughput requirements
- Burst capabilities factored into decision logic
"@
        
        $report | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Detailed report saved: $reportPath" -ForegroundColor Green
        
    } catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
    }
}

# Main execution
try {
    # Check prerequisites
    Write-Host "Checking prerequisites..." -ForegroundColor Yellow
    if (-not (Test-AzureModule)) { 
        Write-Error "Required Azure PowerShell modules are missing. Please install them first."
        exit 1 
    }
    
    if (-not (Connect-ToAzure)) { 
        Write-Error "Failed to connect to Azure. Please check your credentials."
        exit 1 
    }
    
    # Get target DENV subscriptions
    $subscriptions = Get-TargetSubscriptions -TargetNames $TargetSubscriptions
    if ($subscriptions.Count -eq 0) {
        throw "No target DENV/Daikin subscriptions found. Please check subscription names and access permissions."
    }
    
    Write-Host ""
    Write-Host "Starting DENV premium disk analysis..." -ForegroundColor Yellow
    Write-Host "Target subscriptions: $($subscriptions.Count)" -ForegroundColor White
    Write-Host "Utilization threshold: $UtilizationThreshold%" -ForegroundColor White
    Write-Host "Analysis period: $AnalysisDays days" -ForegroundColor White
    
    $allResults = @()
    $totalDisksProcessed = 0
    
    # Analyze each subscription
    foreach ($subscription in $subscriptions) {
        try {
            $premiumDisks = Get-PremiumDisksFromSubscription -SubscriptionId $subscription.Id -SubscriptionName $subscription.Name
            
            if ($premiumDisks.Count -eq 0) {
                Write-Host "  No premium disks found in subscription: $($subscription.Name)" -ForegroundColor Gray
                continue
            }
            
            Write-Host "  Processing $($premiumDisks.Count) premium disks..." -ForegroundColor Gray
            
            foreach ($disk in $premiumDisks) {
                $totalDisksProcessed++
                Write-Progress -Activity "Analyzing Premium Disks" -Status "Processing $($disk.DiskName) ($totalDisksProcessed disks processed)" -PercentComplete (($totalDisksProcessed / ($premiumDisks.Count * $subscriptions.Count)) * 100)
                
                try {
                    # Get real metrics
                    $metrics = Get-RealDiskMetrics -DiskInfo $disk -Days $AnalysisDays -SubscriptionId $subscription.Id
                    
                    # Get recommendation
                    $recommendation = Get-MigrationRecommendation -DiskInfo $disk -Metrics $metrics -Threshold $UtilizationThreshold
                    
                    # Combine all data
                    $result = $disk | Select-Object *, @{
                        Name = "AvgIOPS"
                        Expression = { $metrics.AvgIOPS }
                    }, @{
                        Name = "MaxIOPS"
                        Expression = { $metrics.MaxIOPS }
                    }, @{
                        Name = "AvgReadIOPS"
                        Expression = { $metrics.AvgReadIOPS }
                    }, @{
                        Name = "AvgWriteIOPS"
                        Expression = { $metrics.AvgWriteIOPS }
                    }, @{
                        Name = "AvgThroughputMBps"
                        Expression = { $metrics.AvgThroughputMBps }
                    }, @{
                        Name = "MaxThroughputMBps"
                        Expression = { $metrics.MaxThroughputMBps }
                    }, @{
                        Name = "BaseUtilizationPercent"
                        Expression = { $recommendation.BaseUtilizationPercent }
                    }, @{
                        Name = "BurstUtilizationPercent"
                        Expression = { $recommendation.BurstUtilizationPercent }
                    }, @{
                        Name = "Recommendation"
                        Expression = { $recommendation.Recommendation }
                    }, @{
                        Name = "RecommendedTier"
                        Expression = { $recommendation.RecommendedTier }
                    }, @{
                        Name = "PotentialMonthlySavings"
                        Expression = { $recommendation.PotentialMonthlySavings }
                    }, @{
                        Name = "RecommendationReason"
                        Expression = { $recommendation.RecommendationReason }
                    }, @{
                        Name = "MetricsDataPoints"
                        Expression = { $metrics.DataPointsCollected }
                    }
                    
                    $allResults += $result
                    
                } catch {
                    Write-Warning "Error processing disk $($disk.DiskName): $($_.Exception.Message)"
                }
            }
        } catch {
            Write-Warning "Error processing subscription $($subscription.Name): $($_.Exception.Message)"
        }
    }
    
    Write-Progress -Activity "Analyzing Premium Disks" -Completed
    
    if ($allResults.Count -eq 0) {
        Write-Warning "No premium disks were successfully analyzed."
        exit 1
    }
    
    # Display summary
    Write-Host ""
    Write-Host "Analysis Complete!" -ForegroundColor Green
    Write-Host "=================" -ForegroundColor Green
    Write-Host "Total Premium Disks Analyzed: $($allResults.Count)" -ForegroundColor White
    Write-Host "Migration Candidates: $(($allResults | Where-Object {$_.Recommendation -eq 'Migrate to Standard'}).Count)" -ForegroundColor Yellow
    Write-Host "Deletion Candidates: $(($allResults | Where-Object {$_.Recommendation -like '*Deletion*'}).Count)" -ForegroundColor Red
    Write-Host "Keep Premium: $(($allResults | Where-Object {$_.Recommendation -eq 'Keep Premium'}).Count)" -ForegroundColor Green
    
    $totalSavings = ($allResults | Measure-Object PotentialMonthlySavings -Sum).Sum
    Write-Host "Potential Monthly Savings: $([math]::Round($totalSavings, 2)) USD" -ForegroundColor Green
    Write-Host "Potential Annual Savings: $([math]::Round($totalSavings * 12, 2)) USD" -ForegroundColor Green
    
    # Export results
    Write-Host ""
    Write-Host "Exporting results..." -ForegroundColor Yellow
    Export-Results -Results $allResults -Path $OutputPath -ToExcel:$ExportToExcel
    
    # Display top candidates
    $topMigrationCandidates = $allResults | Where-Object {$_.Recommendation -eq 'Migrate to Standard'} | Sort-Object PotentialMonthlySavings -Descending | Select-Object -First 5
    
    if ($topMigrationCandidates) {
        Write-Host ""
        Write-Host "Top Migration Candidates:" -ForegroundColor Cyan
        $topMigrationCandidates | Format-Table DiskName, DiskTier, RecommendedTier, BaseUtilizationPercent, PotentialMonthlySavings, RecommendationReason -AutoSize
    }
    
    $unattachedDisks = $allResults | Where-Object {$_.Recommendation -like '*Deletion*'}
    if ($unattachedDisks) {
        Write-Host ""
        Write-Host "Unattached Disks (Consider Deletion):" -ForegroundColor Red
        $unattachedDisks | Format-Table DiskName, DiskSizeGB, EstimatedMonthlyCost, CreatedTime -AutoSize
    }
    
    Write-Host ""
    Write-Host "Analysis completed successfully!" -ForegroundColor Green
    Write-Host "Files generated:" -ForegroundColor Gray
    Write-Host "  - Detailed report: $OutputPath.txt" -ForegroundColor Gray
    if ($ExportToExcel) {
        Write-Host "  - Excel workbook: $OutputPath.xlsx" -ForegroundColor Gray
    } else {
        Write-Host "  - CSV data: $OutputPath.csv" -ForegroundColor Gray
    }
    
} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Troubleshooting Steps:" -ForegroundColor Yellow
    Write-Host "1. Ensure Azure PowerShell modules are installed:" -ForegroundColor Gray
    Write-Host "   Install-Module -Name Az -Force -AllowClobber" -ForegroundColor Gray
    Write-Host "2. For Excel export, install ImportExcel module:" -ForegroundColor Gray
    Write-Host "   Install-Module -Name ImportExcel -Force" -ForegroundColor Gray
    Write-Host "3. Verify Azure permissions (Reader role minimum)" -ForegroundColor Gray
    Write-Host "4. Check subscription access and names" -ForegroundColor Gray
    Write-Host "5. Ensure stable internet connection" -ForegroundColor Gray
    
    exit 1
}