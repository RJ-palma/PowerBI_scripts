# Import the Microsoft Power BI Management module
Import-Module -Name MicrosoftPowerBIMgmt

# Function to validate file extension
function IsValidPowerBIFile($file) {
    $validExtensions = @(".pbix", ".pbit")
    return $validExtensions -contains $file.Extension
}

try {
    # Prompt user for the folder path containing Power BI files
    $folderPath = Read-Host "Enter the full path to the folder containing .pbix/.pbit files (i.e. ""C:\Scripts\Downloaded Files\NA1-ESG-PRD-WS-Custom-4284-AirCanada"": )"
    
    # Validate folder path
    if (-not (Test-Path $folderPath)) {
        throw "The specified folder path does not exist: $folderPath"
    }

    # Connect to Power BI service
    Connect-PowerBIServiceAccount

    # Get available workspaces and prompt user to select one
    $workspaces = Get-PowerBIWorkspace -Scope Individual | Select-Object Id, Name
    if ($workspaces.Count -eq 0) {
        throw "No workspaces found for the current user"
    }

    Write-Host "`nAvailable Workspaces:"
    $workspaces | Format-Table -AutoSize | Out-Host
    $workspaceName = Read-Host "Enter the name of the workspace to publish to"

    # Get the selected workspace
    $selectedWorkspace = $workspaces | Where-Object { $_.Name -eq $workspaceName }
    if (-not $selectedWorkspace) {
        throw "Workspace '$workspaceName' not found"
    }

    # Get all .pbix and .pbit files from the specified folder
    $powerBIFiles = Get-ChildItem -Path $folderPath -File | Where-Object { IsValidPowerBIFile($_) }

    if ($powerBIFiles.Count -eq 0) {
        throw "No .pbix or .pbit files found in the specified folder"
    }

    # Publish each valid file
    foreach ($file in $powerBIFiles) {
        Write-Host "Publishing $($file.Name) to workspace '$workspaceName'..."
        
        try {
            # Publish the report with the same name as the file
            New-PowerBIReport -Path $file.FullName `
                            -Name $file.BaseName `
                            -WorkspaceId $selectedWorkspace.Id `
                            -ConflictAction CreateOrOverwrite
            
            Write-Host "Successfully published $($file.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to publish $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "`nPublishing complete!" -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Disconnect from Power BI service
    Disconnect-PowerBIServiceAccount
}