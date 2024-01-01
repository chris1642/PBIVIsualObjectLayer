using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

var basePath = @"C:\Report Backups\Model Backups"; // User-defined base path for Model backups (within Base folder of Steps 1-3)
var modelName = Model.Database.Name; // Retrieve the model name

string GetCurrentDateString() {
    return DateTime.Now.ToString("yyyy-MM-dd"); // Formats the current date as a string
}

var dateFolderPath = Path.Combine(basePath, GetCurrentDateString()); // Generate the path for the date-specific folder
Directory.CreateDirectory(dateFolderPath); // Ensure the directory exists, create it if it doesn't

var filePath = Path.Combine(dateFolderPath, $"{modelName}.csv"); // Construct the full file path for the CSV

string FormatCsvField(string field) {
    if (string.IsNullOrEmpty(field)) {
        return "\"\""; // Return an empty quoted string if the field is null or empty
    }
    return "\"" + field.Replace("\"", "\"\"") + "\""; // Encloses field in double quotes and escapes existing double quotes
}

string header = "Type,Table,Name,FormatString,DisplayFolder,Description,IsHidden,Expression,ModelAsOfDate,ModelName\n"; // CSV header

// Process CalculationGroups
var calculationgroupData = Model.CalculationGroups.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    "", 
    FormatCsvField(m.Name),
    "", 
    "", 
    FormatCsvField(m.Description),
    FormatCsvField(m.IsHidden.ToString()),
    "",  
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process AllColumns
var columnsData = Model.AllColumns.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    FormatCsvField(m.FormatString),
    FormatCsvField(m.DisplayFolder),
    FormatCsvField(m.Description),
    FormatCsvField(m.IsHidden.ToString()),
    "",  
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Measures
var measuresData = Model.AllMeasures.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    FormatCsvField(m.FormatString),
    FormatCsvField(m.DisplayFolder),
    FormatCsvField(m.Description),
    FormatCsvField(m.IsHidden.ToString()),
    FormatCsvField(m.Expression),
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Hierarchies
var hierarchiesData = Model.AllHierarchies.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    "", 
    FormatCsvField(m.DisplayFolder),
    FormatCsvField(m.Description),
    FormatCsvField(m.IsHidden.ToString()),
    "",  
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Levels
var levelsData = Model.AllLevels.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    "", 
    "", 
    FormatCsvField(m.Description),
    "", 
    "", 
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Partitions
var partitionsData = Model.AllPartitions.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    "", 
    "", 
    FormatCsvField(m.Description),
    "", 
    FormatCsvField(m.Expression),
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process KPIs
var kpiData = Model.AllKPIs.Select(m => string.Join(",", new string[]{
    FormatCsvField(m.ObjectType.ToString()),
    FormatCsvField(m.Table.Name),
    FormatCsvField(m.Name),
    "", 
    "", 
    FormatCsvField(m.Description),
    "", 
    FormatCsvField(m.Expression),
    FormatCsvField(GetCurrentDateString()), 
    FormatCsvField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Combine all non-null and non-empty data
var combinedData = new[] { calculationgroupData, columnsData, measuresData, hierarchiesData, levelsData, partitionsData, kpiData }
    .Where(data => data.Any())
    .SelectMany(data => data);

var csvContent = header + string.Join("\n", combinedData);

System.IO.File.WriteAllText(filePath, csvContent); // Write the CSV content to the file