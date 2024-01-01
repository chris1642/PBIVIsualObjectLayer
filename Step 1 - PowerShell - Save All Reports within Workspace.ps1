# Check if MicrosoftPowerBIMgmt is installed, install if not
if (-not (Get-Module -Name MicrosoftPowerBIMgmt -ListAvailable)) {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
}

# Import the module
Import-Module MicrosoftPowerBIMgmt

# Login to Power BI
Login-PowerBI

# Get workspace info from desired Power BI Workspace - Enter Workspace ID below (can be found inside the HTML link of workspace): https://app.powerbi.com/groups/XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX/list?experience=power-bi
$workspaceid = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
$reports = Get-PowerBIReport -WorkspaceId $workspaceid

# Define the base file path
$filepath_base = "C:\Report Backups"

# Create a variable for current date
$date = (Get-Date -UFormat "%Y-%m-%d")

# Create a new folder for the end of week backups
$new_date_folder = Join-Path -Path $filepath_base -ChildPath $date
New-Item -Path $new_date_folder -ItemType Directory -Force

# Export each report in the workspace
foreach ($report in $reports) {
    $filename = $report.Name -replace "[^a-zA-Z0-9]", " " # Clean up the file name
    $filename = "$filename.pbix"
    $filepath = Join-Path -Path $new_date_folder -ChildPath $filename
    Export-PowerBIReport -Id $report.Id -OutFile $filepath
}

Write-Output "Reports exported to $new_date_folder"