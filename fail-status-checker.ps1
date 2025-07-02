# Login to Power BI
Connect-PowerBIServiceAccount

# Get all workspaces
$workspaces = Get-PowerBIWorkspace

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
    Write-Host "Invalid selection. Exiting."
    exit
}

# Get the selected workspace
$selectedWorkspace = $workspaces[$workspaceSelection - 1]

# Get datasets in the selected workspace
$datasets = Get-PowerBIDataset -WorkspaceId $selectedWorkspace.Id

$ctr = 0

if ($datasets.Count -eq 0) {
    Write-Host "No datasets found in workspace '$($selectedWorkspace.Name)'."
} else {
    Write-Host "`nChecking refresh status for datasets in workspace '$($selectedWorkspace.Name)':`n"
    foreach ($dataset in $datasets) {
        # Get the refreshable status for each dataset
        $refreshableStatus = Get-PowerBIDataset -DatasetId $dataset.Id -WorkspaceId $selectedWorkspace.Id
        
        # Only check if the dataset is a semantic model (this assumes models are marked refreshable)
        $isRefreshable = $refreshableStatus.IsRefreshable

        if (-not $isRefreshable) {
            # Only print if the dataset is unrefreshable
            Write-Host "`nDataset '$($dataset.Name)' is " -NoNewline
            Write-Host "**unrefreshable**.`n" -ForegroundColor Yellow

            #uncomment to show error details (does not work)
            # # Check refresh history to find out the reason for failure
            # $refreshHistory = Get-PowerBIRefreshHistory -DatasetId $dataset.Id -WorkspaceId $selectedWorkspace.Id

            # # Sort refresh attempts to find the latest attempt
            # $latestRefresh = $refreshHistory | Sort-Object -Property EndTime -Descending | Select-Object -First 1
            
            # if ($latestRefresh.Status -eq "Failed") {
            #     Write-Host "Last Refresh Attempt: $($latestRefresh.EndTime)"
            #     Write-Host "Failure Reason: $($latestRefresh.ErrorMessage)" -ForegroundColor Red
            # }

            $ctr++


        } else {
            # Uncomment to show refreshable datasets
            # Write-Host "`nDataset '$($dataset.Name)' is **refreshable**.`n"
        }
    }
    Write-Host "`nFinished checking refresh status for datasets in workspace '$($selectedWorkspace.Name)'."
    
    if ($ctr -eq 0) {
        Write-Host "All datasets are refreshable!" -ForegroundColor Green
    }
    else {
        Write-Host "`nThere are " -NoNewline
        Write-Host "'$ctr' " -ForegroundColor Red -NoNewline
        Write-Host "fails."
    }
}