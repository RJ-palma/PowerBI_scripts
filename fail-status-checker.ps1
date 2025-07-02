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
$workspaceSelection = Read-Host "Please enter the number of the workspace you want to check"

# Validate user input
if ($workspaceSelection -lt 1 -or $workspaceSelection -gt $workspaces.Count) {
    Write-Host "Invalid selection. Exiting."
    exit
}

# Get the selected workspace
$selectedWorkspace = $workspaces[$workspaceSelection - 1]

# Get datasets in the selected workspace
$datasets = Get-PowerBIDataset -WorkspaceId $selectedWorkspace.Id

if ($datasets.Count -eq 0) {
    Write-Host "No datasets found in workspace '$($selectedWorkspace.Name)'."
} else {
    # Check refresh history
    foreach ($dataset in $datasets) {
        # Get the refreshable status (type and scheduled status) 
        $refreshableStatus = Get-PowerBIDataset -DatasetId $dataset.Id -WorkspaceId $selectedWorkspace.Id

        # Write the dataset status
        Write-Host "Dataset '$($dataset.Name)' in workspace '$($selectedWorkspace.Name)', Last Refresh: $($refreshableStatus.LastRefresh)," +
                   " Refreshable: $($refreshableStatus.IsRefreshable)"
        
        # Note: You might retrieve activities from Activity Logs instead if available
        # Check Refresh Activity Events for Dataset Refresh Failures
        $refreshActivities = Get-PowerBIActivityEvent -ResourceType Dataset -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)

        foreach ($event in $refreshActivities) {
            if ($event.DatasetId -eq $dataset.Id -and $event.Operation -eq "Refresh" -and $event.Status -eq "Failed") {
                # Log or Notify about the failure
                Write-Host "Dataset '$($dataset.Name)' in workspace '$($selectedWorkspace.Name)' has failed refresh on $($event.EventTime)"
            }
        }
    }
    Write-Host "Finished checking datasets in workspace '$($selectedWorkspace.Name)'."
}