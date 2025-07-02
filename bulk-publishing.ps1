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
    exit# Define log file path
    $LogFile = "C:\Scripts\publish_errors.log"
    
    # Function to log messages
    function Write-Log {
        param ($Message)
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    }
    
    # Check MicrosoftPowerBIMgmt module version
    $module = Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable
    if ($null -eq $module) {
        Write-Host "MicrosoftPowerBIMgmt module not installed. Please run 'Install-Module -Name MicrosoftPowerBIMgmt'." -ForegroundColor Red
        Write-Log "MicrosoftPowerBIMgmt module not installed."
        exit
    }
    else {
        $moduleVersion = $module.Version.ToString()
        Write-Host "Using MicrosoftPowerBIMgmt module version: $moduleVersion" -ForegroundColor Cyan
        Write-Log "Using MicrosoftPowerBIMgmt module version: $moduleVersion"
        if ($moduleVersion -lt "1.2.1026") {
            Write-Host "Warning: Module version is older than 1.2.1026. Consider updating with 'Update-Module -Name MicrosoftPowerBIMgmt'." -ForegroundColor Yellow
            Write-Log "Warning: Module version $moduleVersion is older than 1.2.1026."
        }
    }
    
    # Prompt for Power BI environment
    Write-Host "`nAvailable Power BI environments: Public, USGov, China, Germany, USGovHigh, USGovMil" -ForegroundColor Cyan
    $Environment = Read-Host "Enter the Power BI environment (default: Public)"
    if ([string]::IsNullOrWhiteSpace($Environment)) {
        $Environment = "Public"
    }
    Write-Log "Selected Power BI environment: $Environment"
    
    # Log in to Power BI Service
    try {
        Write-Host "Logging in to Power BI Service (Environment: $Environment)..." -ForegroundColor Cyan
        Login-PowerBI -Environment $Environment
        Write-Host "Login successful." -ForegroundColor Green
        Write-Log "Login successful to $Environment environment."
    }
    catch {
        Write-Host "Error logging in to Power BI: $_" -ForegroundColor Red
        Write-Log "Error logging in to Power BI: $_"
        exit
    }
    
    # Get all workspaces
    Write-Host "`nRetrieving available workspaces..." -ForegroundColor Cyan
    $PBIWorkspaceList = Get-PowerBIWorkspace
    
    # Check if any workspaces are available
    if ($null -eq $PBIWorkspaceList -or $PBIWorkspaceList.Count -eq 0) {
        Write-Host "No workspaces found for the authenticated user in $Environment environment." -ForegroundColor Red
        Write-Log "No workspaces found in $Environment environment."
        exit
    }
    
    # Display the list of workspaces
    Write-Host "`nAvailable Workspaces:" -ForegroundColor Yellow
    Write-Host "--------------------"
    $PBIWorkspaceList | ForEach-Object {
        Write-Host "$($_.name) (ID: $($_.Id))"
    }
    Write-Host "--------------------"
    
    # Prompt user for workspace name
    $WorkspaceName = Read-Host "`nEnter the name of the workspace to publish reports to"
    
    # Find the selected workspace
    $Workspace = $PBIWorkspaceList | Where-Object { $_.name -eq $WorkspaceName }
    
    # Check if the workspace exists
    if ($null -eq $Workspace) {
        Write-Host "Workspace '$WorkspaceName' not found in $Environment environment. Please check the name and try again." -ForegroundColor Red
        Write-Log "Workspace '$WorkspaceName' not found in $Environment environment."
        exit
    }
    Write-Host "Selected workspace: $WorkspaceName (ID: $($Workspace.Id))" -ForegroundColor Cyan
    Write-Log "Selected workspace: $WorkspaceName (ID: $($Workspace.Id))"
    
    # Prompt user for the folder path containing .pbix and .pbit files
    $FolderPath = Read-Host "`nEnter the full path to the folder containing .pbix or .pbit files (e.g., C:\Scripts\PBIXFiles)"
    
    # Check if the folder exists
    if (-not (Test-Path $FolderPath -PathType Container)) {
        Write-Host "Folder '$FolderPath' does not exist. Please check the path and try again." -ForegroundColor Red
        Write-Log "Folder '$FolderPath' does not exist."
        exit
    }
    
    # Get all .pbix and .pbit files in the folder
    $PBIFiles = Get-ChildItem -Path $FolderPath -Filter "*.pbi*" -File
    
    # Check if any .pbix or .pbit files are found
    if ($null -eq $PBIFiles -or $PBIFiles.Count -eq 0) {
        Write-Host "No .pbix or .pbit files found in folder '$FolderPath'." -ForegroundColor Yellow
        Write-Log "No .pbix or .pbit files found in folder '$FolderPath'."
        exit
    }
    
    # Separate .pbix and .pbit files
    $PBIXFiles = $PBIFiles | Where-Object { $_.Extension -eq ".pbix" }
    $PBITFiles = $PBIFiles | Where-Object { $_.Extension -eq ".pbit" }
    
    # Warn about .pbit files
    if ($PBITFiles.Count -gt 0) {
        Write-Host "`nWarning: Found $($PBITFiles.Count) .pbit file(s) in '$FolderPath':" -ForegroundColor Yellow
        $PBITFiles | ForEach-Object {
            Write-Host "- $($_.Name)" -ForegroundColor Yellow
        }
        Write-Host "Power BI Template (.pbit) files cannot be published directly." -ForegroundColor Yellow
        Write-Host "Please open each .pbit file in Power BI Desktop, connect to the required data source, and save as a .pbix file彼此
    
    System: * The response was cut off due to exceeding the token limit. Below is the continuation of the script, completing the artifact and addressing the error further.
    
    <xaiArtifact artifact_id="77f8fc2b-5c16-451d-a6c8-f41c423c0dc4" artifact_version_id="003fd2d0-4abd-43ea-a80e-196a067ae63d" title="Bulk_Publish_PBI_Reports_With_PBIT_Support_Debug.ps1" contentType="text/powershell">
    # Define log file path
    $LogFile = "C:\Scripts\publish_errors.log"
    
    # Function to log messages
    function Write-Log {
        param ($Message)
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    }
    
    # Check MicrosoftPowerBIMgmt module version
    $module = Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable
    if ($null -eq $module) {
        Write-Host "MicrosoftPowerBIMgmt module not installed. Please run 'Install-Module -Name MicrosoftPowerBIMgmt'." -ForegroundColor Red
        Write-Log "MicrosoftPowerBIMgmt module not installed."
        exit
    }
    else {
        $moduleVersion = $module.Version.ToString()
        Write-Host "Using MicrosoftPowerBIMgmt module version: $moduleVersion" -ForegroundColor Cyan
        Write-Log "Using MicrosoftPowerBIMgmt module version: $moduleVersion"
        if ($moduleVersion -lt "1.2.1026") {
            Write-Host "Warning: Module version is older than 1.2.1026. Consider updating with 'Update-Module -Name MicrosoftPowerBIMgmt'." -ForegroundColor Yellow
            Write-Log "Warning: Module version $moduleVersion is older than 1.2.1026."
        }
    }
    
    # Prompt for Power BI environment
    Write-Host "`nAvailable Power BI environments: Public, USGov, China, Germany, USGovHigh, USGovMil" -ForegroundColor Cyan
    $Environment = Read-Host "Enter the Power BI environment (default: Public)"
    if ([string]::IsNullOrWhiteSpace($Environment)) {
        $Environment = "Public"
    }
    Write-Log "Selected Power BI environment: $Environment"
    
    # Log in to Power BI Service
    try {
        Write-Host "Logging in to Power BI Service (Environment: $Environment)..." -ForegroundColor Cyan
        Login-PowerBI -Environment $Environment
        Write-Host "Login successful." -ForegroundColor Green
        Write-Log "Login successful to $Environment environment."
    }
    catch {
        Write-Host "Error logging in to Power BI: $_" -ForegroundColor Red
        Write-Log "Error logging in to Power BI: $_"
        exit
    }
    
    # Get all workspaces
    Write-Host "`nRetrieving available workspaces..." -ForegroundColor Cyan
    $PBIWorkspaceList = Get-PowerBIWorkspace
    
    # Check if any workspaces are available
    if ($null -eq $PBIWorkspaceList -or $PBIWorkspaceList.Count -eq 0) {
        Write-Host "No workspaces found for the authenticated user in $Environment environment." -ForegroundColor Red
        Write-Log "No workspaces found in $Environment environment."
        exit
    }
    
    # Display the list of workspaces
    Write-Host "`nAvailable Workspaces:" -ForegroundColor Yellow
    Write-Host "--------------------"
    $PBIWorkspaceList | ForEach-Object {
        Write-Host "$($_.name) (ID: $($_.Id))"
    }
    Write-Host "--------------------"
    
    # Prompt user for workspace name
    $WorkspaceName = Read-Host "`nEnter the name of the workspace to publish reports to"
    
    # Find the selected workspace
    $Workspace = $PBIWorkspaceList | Where-Object { $_.name -eq $WorkspaceName }
    
    # Check if the workspace exists
    if ($null -eq $Workspace) {
        Write-Host "Workspace '$WorkspaceName' not found in $Environment environment. Please check the name and try again." -ForegroundColor Red
        Write-Log "Workspace '$WorkspaceName' not found in $Environment environment."
        exit
    }
    Write-Host "Selected workspace: $WorkspaceName (ID: $($Workspace.Id))" -ForegroundColor Cyan
    Write-Log "Selected workspace: $WorkspaceName (ID: $($Workspace.Id))"
    
    # Prompt user for the folder path containing .pbix and .pbit files
    $FolderPath = Read-Host "`nEnter the full path to the folder containing .pbix or .pbit files (e.g., C:\Scripts\PBIXFiles)"
    
    # Check if the folder exists
    if (-not (Test-Path $FolderPath -PathType Container)) {
        Write-Host "Folder '$FolderPath' does not exist. Please check the path and try again." -ForegroundColor Red
        Write-Log "Folder '$FolderPath' does not exist."
        exit
    }
    
    # Get all .pbix and .pbit files in the folder
    $PBIFiles = Get-ChildItem -Path $FolderPath -Filter "*.pbi*" -File
    
    # Check if any .pbix or .pbit files are found
    if ($null -eq $PBIFiles -or $PBIFiles.Count -eq 0) {
        Write-Host "No .pbix or .pbit files found in folder '$FolderPath'." -ForegroundColor Yellow
        Write-Log "No .pbix or .pbit files found in folder '$FolderPath'."
        exit
    }
    
    # Separate .pbix and .pbit files
    $PBIXFiles = $PBIFiles | Where-Object { $_.Extension -eq ".pbix" }
    $PBITFiles = $PBIFiles | Where-Object { $_.Extension -eq ".pbit" }
    
    # Warn about .pbit files
    if ($PBITFiles.Count -gt 0) {
        Write-Host "`nWarning: Found $($PBITFiles.Count) .pbit file(s) in '$FolderPath':" -ForegroundColor Yellow
        $PBITFiles | ForEach-Object {
            Write-Host "- $($_.Name)" -ForegroundColor Yellow
        }
        Write-Host "Power BI Template (.pbit) files cannot be published directly." -ForegroundColor Yellow
        Write-Host "Please open each .pbit file in Power BI Desktop, connect to the required data source, and save as a .pbix file in the same folder ('$FolderPath')." -ForegroundColor Yellow
        Write-Host "After converting .pbit files to .pbix, re-run the script to publish them." -ForegroundColor Yellow
        Write-Log "Found $($PBITFiles.Count) .pbit files. User instructed to convert to .pbix."
    }
    
    # Check if any .pbix files are available to publish
    if ($null -eq $PBIXFiles -or $PBIXFiles.Count -eq 0) {
        Write-Host "`nNo .pbix files found to publish in folder '$FolderPath'." -ForegroundColor Yellow
        Write-Log "No .pbix files found to publish in folder '$FolderPath'."
        exit
    }
    
    # Test publishing a single file to catch early errors
    Write-Host "`nTesting publishing with first .pbix file..." -ForegroundColor Cyan
    $TestFile = $PBIXFiles | Select-Object -First 1
    if ($TestFile) {
        $TestReportName = [System.IO.Path]::GetFileNameWithoutExtension($TestFile.Name)
        $TestReportName = $TestReportName -replace '[<>:"/\\|?*]', '_' # Sanitize name
        Write-Host "Test publishing: $TestReportName" -ForegroundColor Yellow
        try {
            New-PowerBIReport -Path $TestFile.FullName -Name $TestReportName -WorkspaceId $Workspace.Id -ConflictAction Overwrite
            Write-Host "Test publish successful: $TestReportName" -ForegroundColor Green
            Write-Log "Test publish successful: $TestReportName"
        }
        catch {
            Write-Host "Test publish failed for '$TestReportName': $_" -ForegroundColor Red
            Write-Log "Test publish failed for '$TestReportName': $_"
            exit
        }
    }
    
    # Publish each .pbix file to the workspace
    Write-Host "`nPublishing $($PBIXFiles.Count) .pbix report(s) to workspace '$WorkspaceName'..." -ForegroundColor Cyan
    Write-Log "Starting to publish $($PBIXFiles.Count) .pbix reports to workspace '$WorkspaceName'."
    $counter = 1
    ForEach ($File in $PBIXFiles) {
        # Validate file existence
        if (-not (Test-Path $File.FullName)) {
            Write-Host "[$counter/$($PBIXFiles.Count)] File not found: $($File.FullName)" -ForegroundColor Red
            Write-Log "File not found: $($File.FullName)"
            $counter++
            continue
        }
    
        # Sanitize report name
        $ReportName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        $ReportName = $ReportName -replace '[<>:"/\\|?*]', '_' # Replace invalid characters
        Write-Host "[$counter/$($PBIXFiles.Count)] Publishing report: $ReportName" -ForegroundColor Yellow
    
        # Re-validate workspace
        try {
            $Workspace = Get-PowerBIWorkspace -Name $WorkspaceName
            if ($null -eq $Workspace) {
                Write-Host "Workspace '$WorkspaceName' no longer exists." -ForegroundColor Red
                Write-Log "Workspace '$WorkspaceName' no longer exists."
                exit
            }
        }
        catch {
            Write-Host "Error validating workspace '$WorkspaceName': $_" -ForegroundColor Red
            Write-Log "Error validating workspace '$WorkspaceName': $_"
            exit
        }
    
        # Publish the report
        try {
            New-PowerBIReport -Path $File.FullName -Name $ReportName -WorkspaceId $Workspace.Id -ConflictAction Overwrite
            Write-Host "Successfully published: $ReportName" -ForegroundColor Green
            Write-Log "Successfully published: $ReportName"
        }
        catch {
            Write-Host "Error publishing report '$ReportName': $_" -ForegroundColor Red
            Write-Log "Error publishing report '$ReportName': $_"
        }
    
        # Delay to avoid rate limiting
        Start-Sleep -Milliseconds 500
        $counter++
    }
    
    Write-Host "`nPublishing process completed for workspace '$WorkspaceName'." -ForegroundColor Green
    Write-Log "Publishing process completed for workspace '$WorkspaceName'."
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