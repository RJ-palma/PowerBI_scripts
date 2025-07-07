# Parameters for reusability
param (
    [string]$TempDir = "C:\Temp",
    [string]$SourceWorkspaceName = "",
    [string]$TargetWorkspaceName = "",
    [string[]]$ReportNames = @(),
    [switch]$CloneAllReports,
    [switch]$CreateNewTargetWorkspace
)

# Ensure the MicrosoftPowerBIMgmt module is installed and imported
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
}
Import-Module MicrosoftPowerBIMgmt

# Centralized logging function
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "Info"
    )
    $color = switch ($Level) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host $Message -ForegroundColor $color
}

# Validate and create temporary directory
$tempFilePath = Join-Path $TempDir "temp.pbix"
try {
    New-Item -Path $TempDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    Write-Log "Temporary directory '$TempDir' is ready." -Level Success
}
catch {
    Write-Log "Failed to create temporary directory '$TempDir': $_" -Level Error
    exit
}

# Function to select workspace
function Select-Workspace {
    param (
        [Parameter(Mandatory)][array]$Workspaces,
        [string]$PromptMessage,
        [string]$PreselectedName = ""
    )
    if ($PreselectedName -and ($selected = $Workspaces | Where-Object { $_.Name -eq $PreselectedName })) {
        return $selected
    }
    Write-Log "`n$PromptMessage"
    $Workspaces | ForEach-Object { $i = 0 } { Write-Log "$($i + 1): $($_.Name)"; $i++ }
    Write-Log "`nEnter the number of the workspace (e.g., 1):"
    $selection = Read-Host
    if ($selection -notmatch '^\d+$' -or [int]$selection -lt 1 -or [int]$selection -gt $Workspaces.Count) {
        Write-Log "Invalid workspace selection: $selection." -Level Error
        exit
    }
    $Workspaces[[int]$selection - 1]
}

# Login to Power BI
try {
    Connect-PowerBIServiceAccount -ErrorAction Stop
    Write-Log "Successfully connected to Power BI." -Level Success
}
catch {
    Write-Log "Failed to connect to Power BI: $_" -Level Error
    exit
}

# Get all workspaces once
try {
    $workspaces = Get-PowerBIWorkspace -ErrorAction Stop | Where-Object { $_.Type -eq "Workspace" }
}
catch {
    Write-Log "Failed to retrieve workspaces: $_" -Level Error
    exit
}

# Select source workspace
$sourceWorkspace = Select-Workspace -Workspaces $workspaces -PromptMessage "Available Workspaces:" -PreselectedName $SourceWorkspaceName
Write-Log "Selected Source Workspace: $($sourceWorkspace.Name)" -Level Success

# Get reports from source workspace
try {
    $reports = Get-PowerBIReport -WorkspaceId $sourceWorkspace.Id -ErrorAction Stop
}
catch {
    Write-Log "Failed to retrieve reports from source workspace '$($sourceWorkspace.Name)': $_" -Level Error
    exit
}

if ($reports.Count -eq 0) {
    Write-Log "No reports found in source workspace '$($sourceWorkspace.Name)'." -Level Warning
    exit
}

# Select reports
$selectedReports = @()
if ($CloneAllReports -or $ReportNames.Count -eq 0) {
    $selectedReports = $reports
} else {
    $selectedReports = $reports | Where-Object { $ReportNames -contains $_.Name }
    if ($selectedReports.Count -eq 0) {
        Write-Log "No matching reports found for names: $($ReportNames -join ', ')." -Level Error
        exit
    }
}
Write-Log "Selected $($selectedReports.Count) report(s) for cloning." -Level Success

# Select or create target workspace(s)
$targetWorkspaces = @()
if ($CreateNewTargetWorkspace -and $TargetWorkspaceName) {
    try {
        $newWorkspace = New-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction Stop
        $targetWorkspaces += $newWorkspace
        Write-Log "Created new target workspace: $TargetWorkspaceName" -Level Success
    }
    catch {
        Write-Log "Failed to create target workspace: $_" -Level Error
        exit
    }
} else {
    $targetWorkspace = Select-Workspace -Workspaces $workspaces -PromptMessage "Available Workspaces:" -PreselectedName $TargetWorkspaceName
    $targetWorkspaces += $targetWorkspace
}
Write-Log "Selected $($targetWorkspaces.Count) target workspace(s) for cloning." -Level Success

# Process each target workspace
foreach ($targetWorkspace in $targetWorkspaces) {
    Write-Log "`n$("=" * 50)"
    Write-Log "Copying Reports to Target Workspace: '$($targetWorkspace.Name)'"
    Write-Log "$("=" * 50)`n"

    $successCount = 0
    $failureCount = 0
    $skippedCount = 0
    $failedReports = @()
    $skippedReports = @()

    foreach ($report in $selectedReports) {
        Write-Log "Processing report: '$($report.Name)'"
        try {
            # Check dataset exportability
            $dataset = Get-PowerBIDataset -WorkspaceId $sourceWorkspace.Id | Where-Object { $_.Id -eq $report.DatasetId }
            if ($dataset -and -not $dataset.IsRefreshable) {
                Write-Log "Report '$($report.Name)' uses a non-refreshable dataset. Copying report only..." -Level Warning
                Copy-PowerBIReport -WorkspaceId $sourceWorkspace.Id -Id $report.Id -TargetWorkspaceId $targetWorkspace.Id -Name $report.Name -ErrorAction Stop
                Write-Log "Successfully copied report '$($report.Name)' (without dataset)." -Level Success
                $successCount++
                $skippedReports += [PSCustomObject]@{ ReportName = $report.Name; Reason = "Non-refreshable dataset (copied report only)" }
                $skippedCount++
                continue
            }

            # Export and import report
            Write-Log "Exporting report '$($report.Name)'..."
            Export-PowerBIReport -WorkspaceId $sourceWorkspace.Id -Id $report.Id -OutFile $tempFilePath -ErrorAction Stop
            Write-Log "Importing report '$($report.Name)' to target workspace..."
            New-PowerBIReport -WorkspaceId $targetWorkspace.Id -Path $tempFilePath -Name $report.Name -ErrorAction Stop
            Write-Log "Successfully copied report '$($report.Name)'." -Level Success
            $successCount++
        }
        catch {
            Write-Log "Failed to process report '$($report.Name)': $_" -Level Error
            $failedReports += [PSCustomObject]@{ ReportName = $report.Name; Reason = $_.Exception.Message }
            $failureCount++
        }
        finally {
            if (Test-Path $tempFilePath) {
                Remove-Item $tempFilePath -Force
                Write-Log "Cleaned up temporary file '$tempFilePath'." -Level Info
            }
        }
    }

    # Summary
    Write-Log "`n$("=" * 50)"
    Write-Log "Copy Operation Summary for '$($targetWorkspace.Name)'"
    Write-Log "$("=" * 50)`n"
    Write-Log "Source Workspace: $($sourceWorkspace.Name)"
    Write-Log "Target Workspace: $($targetWorkspace.Name)"
    Write-Log "Successfully copied: $successCount report(s)" -Level Success
    Write-Log "Failed to copy: $failureCount report(s)" -Level Error
    Write-Log "Skipped (non-exportable): $skippedCount report(s)" -Level Warning

    if ($failedReports) {
        Write-Log "`nFailed Reports:" -Level Error
        $failedReports | ForEach-Object { Write-Log "- '$($_.ReportName)': $($_.Reason)" -Level Error }
    }
    if ($skippedReports) {
        Write-Log "`nSkipped Reports:" -Level Warning
        $skippedReports | ForEach-Object { Write-Log "- '$($_.ReportName)': $($_.Reason)" -Level Warning }
    }
}

# Disconnect from Power BI
Disconnect-PowerBIServiceAccount
Write-Log "Disconnected from Power BI." -Level Success