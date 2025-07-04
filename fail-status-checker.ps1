# Ensure the MicrosoftPowerBIMgmt module is installed and imported
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
}
Import-Module MicrosoftPowerBIMgmt

# Login to Power BI
try {
    Connect-PowerBIServiceAccount -ErrorAction Stop
}
catch {
    Write-Host "Failed to connect to Power BI: $_" -ForegroundColor Red
    exit
}

# Get all workspaces
try {
    $workspaces = Get-PowerBIWorkspace -ErrorAction Stop
}
catch {
    Write-Host "Failed to retrieve workspaces: $_" -ForegroundColor Red
    exit
}

# List available workspaces
Write-Host "Available Workspaces:"
for ($i = 0; $i -lt $workspaces.Count; $i++) {
    Write-Host "$($i + 1): $($workspaces[$i].Name)"
}

# Prompt user for workspace selection
Write-Host "`nPlease enter the numbers of the workspaces you want to check (comma-separated, e.g., 1,2,3) or type 'all' for all workspaces:"
$workspaceSelection = Read-Host

# Handle workspace selection
$selectedWorkspaces = @()
if ($workspaceSelection.Trim().ToLower() -eq "all") {
    $selectedWorkspaces = $workspaces
} else {
    $selections = $workspaceSelection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }
    foreach ($selection in $selections) {
        $index = [int]$selection
        if ($index -ge 1 -and $index -le $workspaces.Count) {
            $selectedWorkspaces += $workspaces[$index - 1]
        } else {
            Write-Host "Invalid selection: $selection. Skipping." -ForegroundColor Red
        }
    }
}

if ($selectedWorkspaces.Count -eq 0) {
    Write-Host "No valid workspaces selected. Exiting." -ForegroundColor Red
    exit
}

# Initialize array to store all failed datasets with workspace info
$allFailedDatasets = @()

# Process each selected workspace
foreach ($workspace in $selectedWorkspaces) {
    Write-Host "`n" + ("=" * 50)
    Write-Host "Checking Workspace: $($workspace.Name)"
    Write-Host ("=" * 50) + "`n"

    # Get datasets in the selected workspace
    try {
        $datasets = Get-PowerBIDataset -WorkspaceId $workspace.Id -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to retrieve datasets for workspace '$($workspace.Name)': $_" -ForegroundColor Red
        continue
    }

    if ($datasets.Count -eq 0) {
        Write-Host "No datasets found in workspace '$($workspace.Name)'." -ForegroundColor Yellow
    } else {
        Write-Host "Checking refresh status for datasets in workspace '$($workspace.Name)':`n"
        foreach ($dataset in $datasets) {
            # Check if the dataset is refreshable
            $isRefreshable = $dataset.IsRefreshable

            if ($isRefreshable) {
                # For refreshable datasets, check refresh history for failures
                try {
                    $apiUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/datasets/$($dataset.Id)/refreshes"
                    $refreshHistory = Invoke-PowerBIRestMethod -Url $apiUrl -Method Get -ErrorAction Stop | ConvertFrom-Json
                }
                catch {
                    Write-Host "Dataset '$($dataset.Name)' refresh history retrieval failed: $_" -ForegroundColor Red
                    $allFailedDatasets += [PSCustomObject]@{
                        WorkspaceName = $workspace.Name
                        DatasetName   = $dataset.Name
                        FailureReason = "Refresh history retrieval failed: $_"
                    }
                    continue
                }

                # Check if refresh history exists
                if ($refreshHistory.value -and $refreshHistory.value.Count -gt 0) {
                    $latestRefresh = $refreshHistory.value | Sort-Object -Property endTime -Descending | Select-Object -First 1
                    if ($latestRefresh.status -eq "Failed") {
                        Write-Host "Dataset '$($dataset.Name)'"
                        Write-Host "is refreshable but " -NoNewline
                        Write-Host "**failed its last refresh**." -ForegroundColor Red
                        Write-Host "Last Refresh Attempt: $($latestRefresh.endTime)"
                        Write-Host "Failure Reason:"
                        Write-Host "$($latestRefresh.serviceExceptionJson)" -ForegroundColor Red
                        $allFailedDatasets += [PSCustomObject]@{
                            WorkspaceName = $workspace.Name
                            DatasetName   = $dataset.Name
                            FailureReason = $latestRefresh.serviceExceptionJson
                        }
                    } else {
                        # Uncomment to show successful refreshes
                        # Write-Host "Dataset '$($dataset.Name)' is refreshable and last refresh status: $($latestRefresh.status)."
                    }
                } else {
                    Write-Host "Dataset '$($dataset.Name)' is refreshable but has no refresh history." -ForegroundColor Yellow
                }
            } else {
                # For unrefreshable datasets
                Write-Host "Dataset '$($dataset.Name)' is **unrefreshable** (e.g., DirectQuery or Live Connection)." -ForegroundColor Yellow
            }
            Write-Host ""  # Add blank line for readability
        }
        Write-Host "Finished checking refresh status for datasets in workspace '$($workspace.Name)'."
    }
}

# Display summary of all failed datasets, grouped by workspace
Write-Host "`n" + ("=" * 50)
Write-Host "Summary of Datasets with Failed Refreshes"
Write-Host ("=" * 50) + "`n"

if ($allFailedDatasets.Count -eq 0) {
    Write-Host "All refreshable datasets across all selected workspaces have no recent failures!" -ForegroundColor Green
} else {
    Write-Host "There are " -NoNewline
    Write-Host "'$($allFailedDatasets.Count)' " -ForegroundColor Red -NoNewline
    Write-Host "datasets with failed refreshes across all selected workspaces."
    Write-Host "Datasets with refresh failures, grouped by workspace:" -ForegroundColor Red

    # Group failed datasets by workspace
    $groupedFailures = $allFailedDatasets | Group-Object -Property WorkspaceName
    foreach ($group in $groupedFailures | Sort-Object -Property Name) {
        Write-Host "`nWorkspace: '$($group.Name)'" -ForegroundColor White
        foreach ($failed in $group.Group) {
            Write-Host "- Dataset: '$($failed.DatasetName)'" -ForegroundColor Red
            Write-Host "  Failure Reason: $($failed.FailureReason)" -ForegroundColor Red
        }
        Write-Host ""  # Add blank line for readability
    }
}