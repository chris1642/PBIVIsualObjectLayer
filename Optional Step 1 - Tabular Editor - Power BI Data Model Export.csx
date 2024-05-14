// Define the base path for the backups
var basePath = @"C:\Power BI Backups";
var addedPath = System.IO.Path.Combine(basePath, "Model Backups");
var modelName = Model.Database.Name; // Retrieve the model name

// Get the current date as a string
var currentDateStr = DateTime.Now.ToString("yyyy-MM-dd");

// Create the directory for today's date if it doesn't exist
var dateFolderPath = System.IO.Path.Combine(addedPath, currentDateStr);
if (!System.IO.Directory.Exists(dateFolderPath))
{
    System.IO.Directory.CreateDirectory(dateFolderPath);
}

// Define the path for the new file
var filePath = System.IO.Path.Combine(dateFolderPath, modelName + ".csv");

// Define the header for the file
var header = "\"Type\",\"Table\",\"Name\",\"FormatString\",\"DisplayFolder\",\"Description\",\"IsHidden\",\"Expression\",\"ModelAsOfDate\",\"ModelName\"";

// Initialize a string builder to collect data
var sb = new System.Text.StringBuilder();
sb.AppendLine(header);

// Function to safely format a field as a string, escape quotes and handle commas
Func<dynamic, string> FormatField = (field) =>
{
    if (string.IsNullOrEmpty(field)) {
        return "\"\""; // Return an empty quoted string if the field is null or empty
    }
    return "\"" + field.Replace("\"", "\"\"") + "\""; 
};

// Add calculation group data to the string builder
foreach (var m in Model.CalculationGroups)
{
    sb.AppendLine(string.Join(",", 
        FormatField("CalculationGroup"),
        FormatField(""),
        FormatField(m.Name),
        FormatField(""),
        FormatField(""),
        FormatField(m.Description),
        FormatField(m.IsHidden.ToString()),
        FormatField(""),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Add columns data to the string builder
foreach (var c in Model.AllColumns)
{
    sb.AppendLine(string.Join(",", 
        FormatField("Column"),
        FormatField(c.Table.Name),
        FormatField(c.Name),
        FormatField(c.FormatString),
        FormatField(c.DisplayFolder),
        FormatField(c.Description),
        FormatField(c.IsHidden.ToString()),
        FormatField(""),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Add calculated columns data to the string builder
foreach (var table in Model.Tables)
{
    foreach (var column in table.Columns.Where(c => c.Type == ColumnType.Calculated))
    {
        sb.AppendLine(string.Join(",", 
            FormatField("CalculatedColumn"),
            FormatField(column.Table.Name),
            FormatField(column.Name),
            FormatField(column.FormatString),
            FormatField(column.DisplayFolder),
            FormatField(column.Description),
            FormatField(column.IsHidden.ToString()),
            FormatField(""),
            FormatField(currentDateStr),
            FormatField(modelName)
        ));
    }
}

// Process Measures
foreach (var am in Model.AllMeasures)
{
    sb.AppendLine(string.Join(",", 
        FormatField("Measure"),
        FormatField(am.Table.Name),
        FormatField(am.Name),
        FormatField(am.FormatString),
        FormatField(am.DisplayFolder),
        FormatField(am.Description),
        FormatField(am.IsHidden.ToString()),
        FormatField(am.Expression),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Process Hierarchies
foreach (var h in Model.AllHierarchies)
{
    sb.AppendLine(string.Join(",", 
        FormatField("Hierarchy"),
        FormatField(h.Table.Name),
        FormatField(h.Name),
        FormatField(""),
        FormatField(h.DisplayFolder),
        FormatField(h.Description),
        FormatField(h.IsHidden.ToString()),
        FormatField(""),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Process Levels
foreach (var l in Model.AllLevels)
{
    sb.AppendLine(string.Join(",", 
        FormatField("Level"),
        FormatField(l.Hierarchy.Table.Name),
        FormatField(l.Name),
        FormatField(""),
        FormatField(""),
        FormatField(l.Description),
        FormatField(""),
        FormatField(""),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Process Partitions
foreach (var p in Model.AllPartitions)
{
    sb.AppendLine(string.Join(",", 
        FormatField("Partition"),
        FormatField(p.Table.Name),
        FormatField(p.Name),
        FormatField(""),
        FormatField(""),
        FormatField(p.Description),
        FormatField(""),
        FormatField(p.Expression),
        FormatField(currentDateStr),
        FormatField(modelName)
    ));
}

// Write the file content to the file
System.IO.File.WriteAllText(filePath, sb.ToString());
 
