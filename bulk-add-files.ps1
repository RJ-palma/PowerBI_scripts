$PbiTools = "C:\Scripts\pbi-tools\pbi-tools.exe"

# Verify the executable exists
if (-not (Test-Path $PbiTools)) {
    Write-Host "Error: pbi-tools.exe not found at $PbiTools"
    exit
}

# # Test PBI Tools version
# Write-Host "Testing PBI Tools..."
# & $PbiTools --version

# # Authenticate with Power BI Service
# Write-Host "Logging in to Power BI Service..."
# try {
#     & $PbiTools login powerbi
# } catch {
#     Write-Host "Login failed: ${_}"
#     exit
# }

# Define workspace ID and output directory
$workspaceId = "me"  # Replace with your workspace ID
$outputDir = "C:\Scripts\Downloaded Files"

# Create output directory if it doesn’t exist
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Get list of reports in the workspace
Write-Host "Fetching reports from workspace $workspaceId..."
try {
    $reportsJson = & $PbiTools powerbi list-reports --workspace-id $workspaceId
    $reports = $reportsJson | ConvertFrom-Json
} catch {
    Write-Host "Failed to list reports: ${_}"
    exit
}

# Loop through each report and download
foreach ($report in $reports) {
    $reportId = $report.id
    $reportName = $report.name -replace '[^\w\-]', '_'  # Replace invalid filename characters
    $outputFile = Join-Path $outputDir "$reportName.pbix"
    Write-Host "Downloading report: $reportName"
    try {
        & $PbiTools powerbi export-report --workspace-id $workspaceId --report-id $reportId --output-path "$outputFile"
        Write-Host "Successfully downloaded: $reportName"
    } catch {
        Write-Host "Failed to download"
    }
}