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
$workspaceSelection = Read-Host "`nPlease enter the number of the workspace you want to check"

# Convert input to integer
$workspaceSelection = [int]$workspaceSelection

# Validate user input
if ($workspaceSelection -lt 1 -or $workspaceSelection -gt $workspaces.Count) {
    Write-Host "Invalid selection. Exiting." -ForegroundColor Red
    exit
}

# Get the selected workspace
$selectedWorkspace = $workspaces[$workspaceSelection - 1]

# Get datasets in the selected workspace
try {
    $datasets = Get-PowerBIDataset -WorkspaceId $selectedWorkspace.Id -ErrorAction Stop
}
catch {
    Write-Host "Failed to retrieve datasets: $_" -ForegroundColor Red
    exit
}

$ctr = 0
$failedDatasets = @()  # Initialize array to store names of datasets with refresh failures

if ($datasets.Count -eq 0) {
    Write-Host "No datasets found in workspace '$($selectedWorkspace.Name)'." -ForegroundColor Yellow
} else {
    Write-Host "`nChecking refresh status for datasets in workspace '$($selectedWorkspace.Name)':`n"
    foreach ($dataset in $datasets) {
        # Check if the dataset is refreshable
        $isRefreshable = $dataset.IsRefreshable

        if ($isRefreshable) {
            # For refreshable datasets, check refresh history for failures
            try {
                $apiUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($selectedWorkspace.Id)/datasets/$($dataset.Id)/refreshes"
                $refreshHistory = Invoke-PowerBIRestMethod -Url $apiUrl -Method Get -ErrorAction Stop | ConvertFrom-Json
            }
            catch {
                Write-Host "`nDataset '$($dataset.Name)' refresh history retrieval failed: $_" -ForegroundColor Red
                $ctr++
                $failedDatasets += $dataset.Name  # Add to failed datasets list
                continue
            }

            # Check if refresh history exists
            if ($refreshHistory.value -and $refreshHistory.value.Count -gt 0) {
                $latestRefresh = $refreshHistory.value | Sort-Object -Property endTime -Descending | Select-Object -First 1
                if ($latestRefresh.status -eq "Failed") {
                    Write-Host "`nDataset '$($dataset.Name)'"
                    Write-Host "is refreshable but " -NoNewline
                    Write-Host "**failed its last refresh**." -ForegroundColor Red
                    Write-Host "Last Refresh Attempt: $($latestRefresh.endTime)"
                    Write-Host "Failure Reason:"
                    Write-Host "$($latestRefresh.serviceExceptionJson)" -ForegroundColor Red
                    $ctr++
                    $failedDatasets += $dataset.Name  # Add to failed datasets list
                } else {
                    # Uncomment to show successful refreshes
                    # Write-Host "`nDataset '$($dataset.Name)' is refreshable and last refresh status: $($latestRefresh.status)."
                }
            } else {
                Write-Host "`nis refreshable but has no refresh history." -ForegroundColor Yellow
                Write-Host "Dataset '$($dataset.Name)' `n"
            }
        } else {
            # For unrefreshable datasets
            Write-Host "`nis **unrefreshable** (e.g., DirectQuery or Live Connection)." -ForegroundColor Yellow
            Write-Host "Dataset '$($dataset.Name)'`n"
            # Note: Unrefreshable datasets typically don’t have refresh history, so skip checking
        }
    }
    Write-Host "`nFinished checking refresh status for datasets in workspace '$($selectedWorkspace.Name)'."
    
    if ($ctr -eq 0) {
        Write-Host "All refreshable datasets have no recent failures!" -ForegroundColor Green
    } else {
        Write-Host "`nThere are " -NoNewline
        Write-Host "'$ctr' " -ForegroundColor Red -NoNewline
        Write-Host "datasets with failed refreshes."
        Write-Host "Datasets with refresh failures:" -ForegroundColor Red
        foreach ($datasetName in $failedDatasets) {
            Write-Host "- $datasetName" -ForegroundColor Red
        }
    }
}