# Ensure the MicrosoftPowerBIMgmt module is installed and imported
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
}
Import-Module MicrosoftPowerBIMgmt

# Login to Power BI
try {
    Connect-PowerBIServiceAccount -ErrorAction Stop
    Write-Host "Successfully connected to Power BI." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Power BI: $_" -ForegroundColor Red
    exit
}

# Set temporary file path (ensure directory exists)
$tempDir = "C:\Temp"
$tempFilePath = "$tempDir\temp.pbix"
try {
    New-Item -Path $tempDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    Write-Host "Temporary directory '$tempDir' is ready." -ForegroundColor Green
}
catch {
    Write-Host "Failed to create temporary directory '$tempDir': $_" -ForegroundColor Red
    exit
}

# Get all workspaces
try {
    $workspaces = Get-PowerBIWorkspace -ErrorAction Stop | Where-Object { $_.Type -eq "Workspace" }
}
catch {
    Write-Host "Failed to retrieve workspaces: $_" -ForegroundColor Red
    exit
}

# List available workspaces
Write-Host "`nAvailable Workspaces:"
for ($i = 0; $i -lt $workspaces.Count; $i++) {
    Write-Host "$($i + 1): $($workspaces[$i].Name)"
}

# Prompt user for source workspace selection
Write-Host "`nPlease enter the number of the source workspace (e.g., 1):"
$sourceSelection = Read-Host

# Validate source workspace selection
if ($sourceSelection -notmatch '^\d+$' -or [int]$sourceSelection -lt 1 -or [int]$sourceSelection -gt $workspaces.Count) {
    Write-Host "Invalid source workspace selection: $sourceSelection. Exiting." -ForegroundColor Red
    exit
}
$sourceWorkspace = $workspaces[[int]$sourceSelection - 1]
Write-Host "Selected Source Workspace: $($sourceWorkspace.Name)" -ForegroundColor Green

# Prompt user for target workspace selection or creation
Write-Host "`nPlease enter the number of the target workspace or type 'new' to create a new workspace:"
$targetSelection = Read-Host

# Handle target workspace selection
if ($targetSelection.Trim().ToLower() -eq "new") {
    Write-Host "Enter the name for the new target workspace:"
    $targetWorkspaceName = Read-Host
    try {
        $targetWorkspace = New-PowerBIWorkspace -Name $targetWorkspaceName -ErrorAction Stop
        Write-Host "Created new target workspace: $targetWorkspaceName" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create target workspace: $_" -ForegroundColor Red
        exit
    }
}
else {
    if ($targetSelection -notmatch '^\d+$' -or [int]$targetSelection -lt 1 -or [int]$targetSelection -gt $workspaces.Count) {
        Write-Host "Invalid target workspace selection: $targetSelection. Exiting." -ForegroundColor Red
        exit
    }
    $targetWorkspace = $workspaces[[int]$targetSelection - 1]
    Write-Host "Selected Target Workspace: $($targetWorkspace.Name)" -ForegroundColor Green
}

# Copy reports and datasets
Write-Host "`n" + ("=" * 50)
Write-Host "Copying Reports and Datasets from '$($sourceWorkspace.Name)' to '$($targetWorkspace.Name)'"
Write-Host ("=" * 50) + "`n"

try {
    $reports = Get-PowerBIReport -WorkspaceId $sourceWorkspace.Id -ErrorAction Stop
}
catch {
    Write-Host "Failed to retrieve reports from source workspace '$($sourceWorkspace.Name)': $_" -ForegroundColor Red
    exit
}

if ($reports.Count -eq 0) {
    Write-Host "No reports found in source workspace '$($sourceWorkspace.Name)'." -ForegroundColor Yellow
    exit
}

$successCount = 0
$failureCount = 0
$skippedCount = 0
$failedReports = @()
$skippedReports = @()

foreach ($report in $reports) {
    Write-Host "`nProcessing report: '$($report.Name)'"
    
    # Get dataset for the report to check if it's exportable
    try {
        $dataset = Get-PowerBIDataset -WorkspaceId $sourceWorkspace.Id | Where-Object { $_.Id -eq $report.DatasetId }
        if ($dataset -and -not $dataset.IsRefreshable) {
            Write-Host "Report '$($report.Name)' uses a non-refreshable dataset (e.g., DirectQuery or Live Connection). Skipping export." -ForegroundColor Yellow
            $skippedReports += [PSCustomObject]@{
                ReportName = $report.Name
                Reason     = "Non-refreshable dataset (e.g., DirectQuery or Live Connection)"
            }
            $skippedCount++
            continue
        }
    }
    catch {
        Write-Host "Failed to check dataset for report '$($report.Name)': $_" -ForegroundColor Yellow
        $skippedReports += [PSCustomObject]@{
            ReportName = $report.Name
            Reason     = "Failed to check dataset: $_"
        }
        $skippedCount++
        continue
    }

    # Export and import report
    try {
        Write-Host "Exporting report '$($report.Name)' to '$tempFilePath'..."
        Export-PowerBIReport -WorkspaceId $sourceWorkspace.Id -Id $report.Id -OutFile $tempFilePath -ErrorAction Stop
        Write-Host "Successfully exported report '$($report.Name)'." -ForegroundColor Green

        Write-Host "Importing report '$($report.Name)' to target workspace..."
        New-PowerBIReport -WorkspaceId $targetWorkspace.Id -Path $tempFilePath -Name $report.Name -ErrorAction Stop
        Write-Host "Successfully copied report '$($report.Name)' to target workspace." -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "Failed to process report '$($report.Name)': $_" -ForegroundColor Red
        $failedReports += [PSCustomObject]@{
            ReportName = $report.Name
            Reason     = $_.Exception.Message
        }
        $failureCount++
    }
    finally {
        # Clean up temporary file
        if (Test-Path $tempFilePath) {
            Remove-Item $tempFilePath -Force
            Write-Host "Cleaned up temporary file '$tempFilePath'." -ForegroundColor Gray
        }
    }
}

# Summary
Write-Host "`n" + ("=" * 50)
Write-Host "Copy Operation Summary"
Write-Host ("=" * 50) + "`n"
Write-Host "Source Workspace: $($sourceWorkspace.Name)"
Write-Host "Target Workspace: $($targetWorkspace.Name)"
Write-Host "Successfully copied: $successCount report(s)" -ForegroundColor Green
Write-Host "Failed to copy: $failureCount report(s)" -ForegroundColor Red
Write-Host "Skipped (non-exportable): $skippedCount report(s)" -ForegroundColor Yellow

# List failed reports
if ($failedReports.Count -gt 0) {
    Write-Host "`nFailed Reports:" -ForegroundColor Red
    foreach ($failed in $failedReports) {
        Write-Host "- '$($failed.ReportName)': $($failed.Reason)" -ForegroundColor Red
    }
}

# List skipped reports
if ($skippedReports.Count -gt 0) {
    Write-Host "`nSkipped Reports:" -ForegroundColor Yellow
    foreach ($skipped in $skippedReports) {
        Write-Host "- '$($skipped.ReportName)': $($skipped.Reason)" -ForegroundColor Yellow
    }
}

# Disconnect from Power BI
Disconnect-PowerBIServiceAccount
Write-Host "`nDisconnected from Power BI." -ForegroundColor Green