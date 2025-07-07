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
$WorkspaceName = Read-Host "`nEnter the name of the workspace to publish reports to"

# Find the selected workspace
$Workspace = $PBIWorkspaceList | Where-Object { $_.name -eq $WorkspaceName }

# Check if the workspace exists
if ($null -eq $Workspace) {
    Write-Host "Workspace '$WorkspaceName' not found. Please check the name and try again." -ForegroundColor Red
    exit
}

# Prompt user for the folder path containing .pbix files
$FolderPath = Read-Host "`nEnter the full path to the folder containing .pbix files (e.g., C:\Scripts\PBIXFiles)"

# Check if the folder exists
if (-not (Test-Path $FolderPath -PathType Container)) {
    Write-Host "Folder '$FolderPath' does not exist. Please check the path and try again." -ForegroundColor Red
    exit
}

# Get all .pbix files in the folder
$PBIXFiles = Get-ChildItem -Path $FolderPath -Filter "*.pbix" -File

# Check if any .pbix files are found
if ($null -eq $PBIXFiles -or $PBIXFiles.Count -eq 0) {
    Write-Host "No .pbix files found in folder '$FolderPath'." -ForegroundColor Yellow
    exit
}

# Publish each .pbix file to the workspace
Write-Host "`nPublishing $($PBIXFiles.Count) report(s) to workspace '$WorkspaceName'..." -ForegroundColor Cyan
$counter = 1
ForEach ($File in $PBIXFiles) {
    $ReportName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    Write-Host "[$counter/$($PBIXFiles.Count)] Publishing report: $ReportName" -ForegroundColor Yellow

    try {
        New-PowerBIReport -Path $File.FullName -Name $ReportName -WorkspaceId $Workspace.Id -ConflictAction Overwrite
        Write-Host "Successfully published: $ReportName" -ForegroundColor Green
    }
    catch {
        Write-Host "Error publishing report '$ReportName': $_" -ForegroundColor Red
    }
    $counter++
}

Write-Host "`nPublishing process completed for workspace '$WorkspaceName'." -ForegroundColor Green