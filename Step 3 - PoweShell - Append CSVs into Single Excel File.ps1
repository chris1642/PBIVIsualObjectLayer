# Define the same base folder path defined in step 1 and 2
$baseFolderPath = "C:\Report Backups"

# Install the ImportExcel module if not already installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser
}

# Create a variable for the current date
$currentDate = Get-Date -UFormat "%Y-%m-%d"

# Create a new folder for the current date
$currentDateFolder = Join-Path -Path $baseFolderPath -ChildPath $currentDate
New-Item -Path $currentDateFolder -ItemType Directory -Force

# Define the parent folder path
$parentFolderPath = Split-Path -Path $baseFolderPath -Parent

# Define the output Excel file path in the parent folder
$outputExcelFile = Join-Path -Path $parentFolderPath -ChildPath "VisualObjectLayer.xlsx"

# Check if the Excel file already exists
if (Test-Path -Path $outputExcelFile) {
    $excelExists = $true
} else {
    $excelExists = $false
}

foreach ($csvFile in (Get-ChildItem -Path $currentDateFolder -Filter *.csv)) {
    # Get the base name of the file (without extension) for the worksheet name
    $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.FullName)
    
    # Import the CSV file
    $csvData = Import-Csv -Path $csvFile.FullName

    if ($excelExists) {
        # Append data to the existing Excel file
        $csvData | Export-Excel -Path $outputExcelFile -WorksheetName $worksheetName -AutoNameRange -Append
    } else {
        # Create a new Excel file with the data
        $csvData | Export-Excel -Path $outputExcelFile -WorksheetName $worksheetName -AutoNameRange
        $excelExists = $true
    }
}

if ($excelExists) {
    Write-Output "CSV files appended to $outputExcelFile"
} else {
    Write-Output "CSV files combined and exported to $outputExcelFile"
}