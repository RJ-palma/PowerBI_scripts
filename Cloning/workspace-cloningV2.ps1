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

# Prompt user to clone entire workspace or select specific reports
Write-Host "`nDo you want to clone the entire workspace or select specific reports? (Enter 'all' or 'select'):"
$cloneOption = Read-Host

# Get reports from source workspace
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

# Handle report selection
$selectedReports = @()
if ($cloneOption.Trim().ToLower() -eq "select") {
    Write-Host "`nAvailable Reports in '$($sourceWorkspace.Name)':"
    for ($i = 0; $i -lt $reports.Count; $i++) {
        Write-Host "$($i + 1): $($reports[$i].Name)"
    }
    Write-Host "`nPlease enter the numbers of the reports to clone (comma-separated, e.g., 1,3,5) or 'all' for all reports:"
    $reportSelection = Read-Host
    if ($reportSelection.Trim().ToLower() -eq "all") {
        $selectedReports = $reports
    } else {
        $selections = $reportSelection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
        foreach ($selection in $selections) {
            $index = [int]$selection
            if ($index -ge 1 -and $index -le $reports.Count) {
                $selectedReports += $reports[$index - 1]
            } else {
                Write-Host "Invalid report selection: $selection. Skipping." -ForegroundColor Yellow
            }
        }
        if ($selectedReports.Count -eq 0) {
            Write-Host "No valid reports selected. Exiting." -ForegroundColor Red
            exit
        }
    }
} else {
    $selectedReports = $reports
}
Write-Host "Selected $($selectedReports.Count) report(s) for cloning." -ForegroundColor Green

# Prompt user for target workspace(s) selection or creation
Write-Host "`nDo you want to clone into a new workspace or select existing workspace(s)? (Enter 'new' or 'select'):"
$targetOption = Read-Host

$targetWorkspaces = @()
if ($targetOption.Trim().ToLower() -eq "new") {
    Write-Host "Enter the name for the new target workspace:"
    $targetWorkspaceName = Read-Host
    try {
        $newWorkspace = New-PowerBIWorkspace -Name $targetWorkspaceName -ErrorAction Stop
        $targetWorkspaces += $newWorkspace
        Write-Host "Created new target workspace: $targetWorkspaceName" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create target workspace: $_" -ForegroundColor Red
        exit
    }
} else {

    # List available workspaces
    Write-Host "`nAvailable Workspaces:"
    for ($i = 0; $i -lt $workspaces.Count; $i++) {
        Write-Host "$($i + 1): $($workspaces[$i].Name)"
    }

    Write-Host "`nPlease enter the numbers of the target workspaces (comma-separated, e.g., 1,3,5):"
    $targetSelection = Read-Host
    $selections = $targetSelection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
    foreach ($selection in $selections) {
        $index = [int]$selection
        if ($index -ge 1 -and $index -le $workspaces.Count) {
            $targetWorkspaces += $workspaces[$index - 1]
        } else {
            Write-Host "Invalid target workspace selection: $selection. Skipping." -ForegroundColor Yellow
        }
    }
    if ($targetWorkspaces.Count -eq 0) {
        Write-Host "No valid target workspaces selected. Exiting." -ForegroundColor Red
        exit
    }
}
Write-Host "Selected $($targetWorkspaces.Count) target workspace(s) for cloning." -ForegroundColor Green

# Process each target workspace
foreach ($targetWorkspace in $targetWorkspaces) {
    Write-Host "`n" + ("=" * 50)
    Write-Host "Copying Reports to Target Workspace: '$($targetWorkspace.Name)'"
    Write-Host ("=" * 50) + "`n"

    $successCount = 0
    $failureCount = 0
    $skippedCount = 0
    $failedReports = @()
    $skippedReports = @()

    foreach ($report in $selectedReports) {
        Write-Host "`nProcessing report: '$($report.Name)'"
        
        # Get dataset for the report to check if it's exportable
        try {
            $dataset = Get-PowerBIDataset -WorkspaceId $sourceWorkspace.Id | Where-Object { $_.Id -eq $report.DatasetId }
            if ($dataset -and -not $dataset.IsRefreshable) {
                Write-Host "Report '$($report.Name)' uses a non-refreshable dataset (e.g., DirectQuery or Live Connection). Attempting to copy report only..." -ForegroundColor Yellow
                try {
                    Copy-PowerBIReport -WorkspaceId $sourceWorkspace.Id -Id $report.Id -TargetWorkspaceId $targetWorkspace.Id -Name $report.Name -ErrorAction Stop
                    Write-Host "Successfully copied report '$($report.Name)' (without dataset) to target workspace." -ForegroundColor Green
                    $successCount++
                    $skippedReports += [PSCustomObject]@{
                        ReportName = $report.Name
                        Reason     = "Non-refreshable dataset (copied report only)"
                    }
                    $skippedCount++
                    continue
                }
                catch {
                    Write-Host "Failed to copy report '$($report.Name)' (without dataset): $_" -ForegroundColor Red
                    $failedReports += [PSCustomObject]@{
                        ReportName = $report.Name
                        Reason     = "Copy report failed: $_"
                    }
                    $failureCount++
                    continue
                }
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

    # Summary for current target workspace
    Write-Host "`n" + ("=" * 50)
    Write-Host "Copy Operation Summary for '$($targetWorkspace.Name)'"
    Write-Host ("=" * 50) + "`n"
    Write-Host "Source Workspace: $($sourceWorkspace.Name)"
    Write-Host "Target Workspace: $($targetWorkspace.Name)"
    Write-Host "Successfully copied: $successCount report(s)" -ForegroundColor Green
    Write-Host "Failed to copy: $failureCount report(s)" -ForegroundColor Red
    Write-Host "Skipped (non-exportable): $skippedCount report(s)" -ForegroundColor Yellow

    if ($failedReports.Count -gt 0) {
        Write-Host "`nFailed Reports:" -ForegroundColor Red
        foreach ($failed in $failedReports) {
            Write-Host "- '$($failed.ReportName)': $($failed.Reason)" -ForegroundColor Red
        }
    }

    if ($skippedReports.Count -gt 0) {
        Write-Host "`nSkipped Reports:" -ForegroundColor Yellow
        foreach ($skipped in $skippedReports) {
            Write-Host "- '$($skipped.ReportName)': $($skipped.Reason)" -ForegroundColor Yellow
        }
    }
}

# Disconnect from Power BI
Disconnect-PowerBIServiceAccount
Write-Host "`nDisconnected from Power BI." -ForegroundColor Green