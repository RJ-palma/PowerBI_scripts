# Log in to Power BI Service
try {
    Write-Host "Logging in to Power BI Service..." -ForegroundColor Cyan
    Login-PowerBI -Environment Public
    Write-Host "Login successful." -ForegroundColor Green
}
catch {
    Write-Host "Error logging in to Power BI: $_" -ForegroundColor Red
    exit
}

# Get all workspaces
Write-Host "`nRetrieving available workspaces..." -ForegroundColor Cyan
$PBIWorkspaceList = Get-PowerBIWorkspace

# Check if any workspaces are available
if ($null -eq $PBIWorkspaceList -or $PBIWorkspaceList.Count -eq 0) {
    Write-Host "No workspaces found for the authenticated user." -ForegroundColor Red
    exit
}

# Display the list of workspaces
Write-Host "`nAvailable Workspaces:" -ForegroundColor Yellow
Write-Host "--------------------"
$PBIWorkspaceList | ForEach-Object {
    Write-Host $_.name
}
Write-Host "--------------------"

# Prompt user for workspace name
$WorkspaceName = Read-Host "`nEnter the name of the workspace to download reports from"

# Find the selected workspace
$Workspace = $PBIWorkspaceList | Where-Object { $_.name -eq $WorkspaceName }

# Check if the workspace exists
if ($null -eq $Workspace) {
    Write-Host "Workspace '$WorkspaceName' not found. Please check the name and try again." -ForegroundColor Red
    exit
}

# Create a folder for the workspace
$FolderName = "C:\Scripts\Downloaded Files\$($Workspace.name)"
Write-Host "`nCreating folder: $FolderName" -ForegroundColor Cyan
New-Item -Path $FolderName -ItemType Directory -Force | Out-Null

# Get all reports in the workspace
Write-Host "Retrieving reports from workspace '$WorkspaceName'..." -ForegroundColor Cyan
$PBIReports = Get-PowerBIReport -WorkspaceId $Workspace.Id

# Check if any reports are available
if ($null -eq $PBIReports -or $PBIReports.Count -eq 0) {
    Write-Host "No reports found in workspace '$WorkspaceName'." -ForegroundColor Yellow
    exit
}

# Loop through each report with progress feedback
Write-Host "Downloading $($PBIReports.Count) report(s)..." -ForegroundColor Cyan
$counter = 1
ForEach ($Report in $PBIReports) {
    Write-Host "`n[$counter/$($PBIReports.Count)] Downloading report: $($Report.name)" -ForegroundColor Yellow
    $OutputFile = "$FolderName\$($Report.name).pbix"

    # If the file exists, delete it
    if (Test-Path $OutputFile) {
        Write-Host "Existing file found at $OutputFile. Deleting..." -ForegroundColor Gray
        Remove-Item $OutputFile -Force
    }

    # Export the report
    try {
        Export-PowerBIReport -WorkspaceId $Workspace.Id -Id $Report.Id -OutFile $OutputFile
        Write-Host "Successfully downloaded: $($Report.name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error downloading report '$($Report.name)': $_" -ForegroundColor Red
    }
    $counter++
}

Write-Host "`nDownload process completed for workspace '$WorkspaceName'." -ForegroundColor Green