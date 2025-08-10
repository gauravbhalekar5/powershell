## Azure Storage FinOps Analyzer

This PowerShell script analyzes Azure Storage accounts across all subscriptions to find cost optimization opportunities (FinOps-aligned). It gathers usage metrics and costs, and produces console output plus optional CSV, Excel (with charts), and a LaTeX/PDF report.

### Prerequisites
- PowerShell 7+ (`pwsh`)
- Azure PowerShell modules:
  - `Az.Accounts`, `Az.Storage`, `Az.Monitor`, `Az.Consumption`
  - Install if missing: `Install-Module Az -Scope CurrentUser`
- Optional for Excel export: `ImportExcel` module
  - `Install-Module ImportExcel -Scope CurrentUser`
- Optional for PDF: `latexmk` (from TeX Live or MikTeX)

### Usage
Run from any shell that has PowerShell 7+ available. Outputs are written to `./out` by default.

```bash
pwsh -File /workspace/azure-storage-finops/StorageCostAnalysis.ps1 \
  -Days 30 \
  -ThresholdGB 10 \
  -ExportCsv \
  -ExportExcel \
  -ExportPdf
```

#### Parameters
- `-Days <int>`: Analysis period in days (default: 30)
- `-ThresholdGB <int>`: Low usage threshold (default: 10 GB)
- `-ExportCsv`: Export detailed report to CSV
- `-ExportExcel`: Export to Excel with charts (requires `ImportExcel`)
- `-ExportPdf`: Generate LaTeX and, if `latexmk` is available, compile to PDF
- `-OutputDir <path>`: Output directory (default: `./out`)

### Notes
- Cost estimates use rough 2025 US East prices; validate for your region using the Azure Pricing Calculator.
- Consumption API is fetched once per subscription to improve performance; metrics are retrieved per account.
- Lifecycle policy example is included as guidance; adjust to your governance and Az module version.

### Advanced analyzer
The advanced script provides multi-region pricing, security checks, richer recommendations, and additional export options.

```bash
pwsh -File /workspace/azure-storage-finops/AdvancedStorageCostOptimization.ps1 \
  -Days 30 \
  -ThresholdGB 10 \
  -ExportCsv -ExportExcel -ExportHtml -ExportPdf \
  -OutputDir /workspace/azure-storage-finops/reports \
  -EmailReport you@example.com
```

Notes:
- Email sending uses `Send-MailMessage` and requires environment variables: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`.
- Auto remediation writes commands to `Remediation_Actions.ps1` for review instead of executing by default.
- Region defaults to the account location, but you can override with `-Region <azure-region>`.