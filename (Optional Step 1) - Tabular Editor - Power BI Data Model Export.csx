using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

var basePath = @"C:\Power BI Backups"; // User-defined base path

var addedPath = Path.Combine(basePath, "Model Backups");
var modelName = Model.Database.Name; // Retrieve the model name

string GetCurrentDateString() {
    return DateTime.Now.ToString("yyyy-MM-dd"); // Formats the current date as a string
}

var dateFolderPath = Path.Combine(addedPath, GetCurrentDateString()); // Generate the path for the date-specific folder
Directory.CreateDirectory(dateFolderPath); // Ensure the directory exists, create it if it doesn't

var filePath = Path.Combine(dateFolderPath, $"{modelName}.txt"); // Construct the full file path for the text file

string FormatField(string field) {
    if (string.IsNullOrEmpty(field)) {
        return ""; // Return an empty string if the field is null or empty
    }
    return field; // For text files, you might not need to enclose fields in quotes, but you can adjust this as needed
}

// Update the header to use tabs
string header = "Type\tTable\tName\tFormatString\tDisplayFolder\tDescription\tIsHidden\tExpression\tModelAsOfDate\tModelName\n";

// Process CalculationGroups
var calculationgroupData = Model.CalculationGroups.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    "", 
    FormatField(m.Name),
    "", 
    "", 
    FormatField(m.Description),
    FormatField(m.IsHidden.ToString()),
    "",  
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process AllColumns
var columnsData = Model.AllColumns.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    FormatField(m.FormatString),
    FormatField(m.DisplayFolder),
    FormatField(m.Description),
    FormatField(m.IsHidden.ToString()),
    "",  
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Measures
var measuresData = Model.AllMeasures.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    FormatField(m.FormatString),
    FormatField(m.DisplayFolder),
    FormatField(m.Description),
    FormatField(m.IsHidden.ToString()),
    FormatField(m.Expression),
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Hierarchies
var hierarchiesData = Model.AllHierarchies.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    "", 
    FormatField(m.DisplayFolder),
    FormatField(m.Description),
    FormatField(m.IsHidden.ToString()),
    "",  
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Levels
var levelsData = Model.AllLevels.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    "", 
    "", 
    FormatField(m.Description),
    "", 
    "", 
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process Partitions
var partitionsData = Model.AllPartitions.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    "", 
    "", 
    FormatField(m.Description),
    "", 
    FormatField(m.Expression),
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Process KPIs
var kpiData = Model.AllKPIs.Select(m => string.Join("\t", new string[]{
    FormatField(m.ObjectType.ToString()),
    FormatField(m.Table.Name),
    FormatField(m.Name),
    "", 
    "", 
    FormatField(m.Description),
    "", 
    FormatField(m.Expression),
    FormatField(GetCurrentDateString()), 
    FormatField(modelName) 
})).Where(line => !string.IsNullOrWhiteSpace(line)).ToList();

// Combine all non-null and non-empty data
var combinedData = new[] { calculationgroupData, columnsData, measuresData, hierarchiesData, levelsData, partitionsData, kpiData }
    .Where(data => data.Any())
    .SelectMany(data => data);

var textContent = header + string.Join("\n", combinedData);

System.IO.File.WriteAllText(filePath, textContent); // Write the text content to the file
