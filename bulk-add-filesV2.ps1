#Create a New Folder
# New-Item -Path 'D:\AllPBIFiles\' -ItemType Directory

<#
  This script requires the MicrosoftPowerBIMgmt module.
  Run the following commands to install it if not already installed: 
  Find-Module -Name PowerShellGet
  Install-Module -Name MicrosoftPowerBIMgmt 
#>
 
<#If the login step to PBI fails, try running the following line first.
 [Net.ServicePointManager]:SecurityProtocol = [Net.SecurityProtocolType]::Tsl12
#>

#Log in to Power BI Service
Login-PowerBI  -Environment Public

### Function to download all reports of all workspaces 
### Will explode your pc - be warned - 

# $PBIWorkspaceList = Get-PowerBIWorkspace #Get All Workspace Name

# #Loop through the Workspace:
# ForEach($Workspace in $PBIWorkspaceList)
# {
#  $FolderName = "C:\Scripts\Downloaded Files"+ $Workspace.name
#  New-Item -Path $FolderName -ItemType Directory

# 	$PBIReports = Get-PowerBIReport -WorkspaceId $Workspace.Id 
	  
# 		#Loop through the reports:
# 		ForEach($Report in $PBIReports)
# 		{
# 		 Write-Host $Report.name 		 
# 		 $OutputFile = $FolderName+ "\" +  $Report.name + ".pbix"
		 
# 		 # If the file exists, delete it first; otherwise, the Export-PowerBIReport will fail.
# 		 if (Test-Path $OutputFile)
# 		 {
# 			 Remove-Item $OutputFile
# 		 }
		 
# 		 Export-PowerBIReport -WorkspaceId $PBIWorkspace.ID -Id $Report.ID -OutFile $OutputFile
# 		}
 
# }

# Define the target workspace name or id
$WorkspaceName = "me"

# Get the specific workspace by name
$Workspace = Get-PowerBIWorkspace -Name $WorkspaceName

# Check if the workspace exists
if ($null -eq $Workspace) {
    Write-Host "Workspace '$WorkspaceName' not found."
    exit
}

# Create a folder for the workspace
$FolderName = "C:\Scripts\Downloaded Files\$($Workspace.name)"
New-Item -Path $FolderName -ItemType Directory -Force

# Get all reports in the workspace
$PBIReports = Get-PowerBIReport -WorkspaceId $Workspace.Id

# Loop through each report
ForEach ($Report in $PBIReports) {
    Write-Host $Report.name
    $OutputFile = "$FolderName\$($Report.name).pbix"

    # If the file exists, delete it
    if (Test-Path $OutputFile) {
        Remove-Item $OutputFile
    }

    # Export the report
    Export-PowerBIReport -WorkspaceId $Workspace.Id -Id $Report.Id -OutFile $OutputFile
}
