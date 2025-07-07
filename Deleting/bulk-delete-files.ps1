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

# Prompt user for workspace selection
Write-Host "`nPlease enter the number of the workspace (e.g., 1):"
$selection = Read-Host

# Validate workspace selection
if ($selection -notmatch '^\d+$' -or [int]$selection -lt 1 -or [int]$selection -gt $workspaces.Count) {
    Write-Host "Invalid workspace selection: $selection. Exiting." -ForegroundColor Red
    exit
}
$selectedWorkspace = $workspaces[[int]$selection - 1]
Write-Host "Selected Workspace: $($selectedWorkspace.Name)" -ForegroundColor Green

# Prompt user for deletion option
Write-Host "`nDo you want to delete the entire workspace or select specific reports to delete? (Enter 'entire' or 'select'):"
$deleteOption = Read-Host

if ($deleteOption.Trim().ToLower() -eq "entire") {
    # First confirmation prompt
    Write-Host "`nWARNING: You are about to delete the entire workspace '$($selectedWorkspace.Name)'. This will permanently remove all reports, datasets, and other content in the workspace."
    Write-Host "Are you sure you want to proceed? (Enter 'yes' to continue, any other input to cancel):"
    $firstConfirm = Read-Host
    if ($firstConfirm.Trim().ToLower() -ne "yes") {
        Write-Host "Workspace deletion cancelled." -ForegroundColor Yellow
        Disconnect-PowerBIServiceAccount
        exit
    }

    # Second confirmation prompt
    Write-Host "`nFINAL CONFIRMATION: This action cannot be undone. Confirm deletion of workspace '$($selectedWorkspace.Name)'? (Enter 'DELETE' to confirm, any other input to cancel):"
    $secondConfirm = Read-Host
    if ($secondConfirm.Trim().ToLower() -ne "delete") {
        Write-Host "Workspace deletion cancelled." -ForegroundColor Yellow
        Disconnect-PowerBIServiceAccount
        exit
    }

    # Delete the entire workspace
    try {
        Remove-PowerBIWorkspace -Id $selectedWorkspace.Id -ErrorAction Stop
        Write-Host "Successfully deleted workspace '$($selectedWorkspace.Name)'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to delete workspace '$($selectedWorkspace.Name)': $_" -ForegroundColor Red
    }
}
else {
    # Get reports from selected workspace
    try {
        $reports = Get-PowerBIReport -WorkspaceId $selectedWorkspace.Id -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to retrieve reports from workspace '$($selectedWorkspace.Name)': $_" -ForegroundColor Red
        Disconnect-PowerBIServiceAccount
        exit
    }

    if ($reports.Count -eq 0) {
        Write-Host "No reports found in workspace '$($selectedWorkspace.Name)'." -ForegroundColor Yellow
        Disconnect-PowerBIServiceAccount
        exit
    }

    # List available reports
    Write-Host "`nAvailable Reports in '$($selectedWorkspace.Name)':"
    for ($i = 0; $i -lt $reports.Count; $i++) {
        Write-Host "$($i + 1): $($reports[$i].Name)"
    }

    # Prompt user for report selection
    Write-Host "`nPlease enter the numbers of the reports to delete (comma-separated, e.g., 1,3,5):"
    $reportSelection = Read-Host
    $selections = $reportSelection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }

    $selectedReports = @()
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
        Disconnect-PowerBIServiceAccount
        exit
    }

    # Process report deletions
    $successCount = 0
    $failureCount = 0
    $failedReports = @()

    foreach ($report in $selectedReports) {
        Write-Host "`nProcessing report: '$($report.Name)'"
        try {
            Remove-PowerBIReport -WorkspaceId $selectedWorkspace.Id -Id $report.Id -ErrorAction Stop
            Write-Host "Successfully deleted report '$($report.Name)'." -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "Failed to delete report '$($report.Name)': $_" -ForegroundColor Red
            $failedReports += [PSCustomObject]@{
                ReportName = $report.Name
                Reason     = $_.Exception.Message
            }
            $failureCount++
        }
    }

    # Summary of report deletions
    Write-Host "`n" + ("=" * 50)
    Write-Host "Deletion Summary for Workspace '$($selectedWorkspace.Name)'"
    Write-Host ("=" * 50) + "`n"
    Write-Host "Successfully deleted: $successCount report(s)" -ForegroundColor Green
    Write-Host "Failed to delete: $failureCount report(s)" -ForegroundColor Red

    if ($failedReports.Count -gt 0) {
        Write-Host "`nFailed Reports:" -ForegroundColor Red
        foreach ($failed in $failedReports) {
            Write-Host "- '$($failed.ReportName)': $($failed.Reason)" -ForegroundColor Red
        }
    }
}

# Disconnect from Power BI
Disconnect-PowerBIServiceAccount
Write-Host "`nDisconnected from Power BI." -ForegroundColor Green