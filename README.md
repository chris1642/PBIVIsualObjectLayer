Recommended: use a Base folder that is cloud-based (such as SharePoint) - this will make it easy to leverage within a Power BI semantic model and automatic refreshes. 

1. PowerShell - Define Workspace ID and Base Folder path - run within PowerShell
2. Tabular Editor 3 - Define Base Folder path and run C# script within Tabular Editor 3
3. PowerShell - Define Base Folder path and run within PowerShell
4. Tabular Editor 3 (Optional) - Define Base Model Backup Folder, Connect to Semantic Model within PBI, and run C# script
5. PowerShell (Optional) - Define Base Model Backup Folder, Excel file path, and Excel File name and run within PowerShell


Once you have all 5 set up....downloading the PBIT will begin a refresh that is connected to C:/Report Backups and a Usage Metrics DirectQuery. When the initial refresh fails, click Transform Data -> Data Source settings and change the 3 connections to your Visual Object Layer Excel file, Model Detail excel file, and the Usage Metrics Report from the same Workspace you have downloaded reports from. 

Once the refresh is completed, the measures and relationships are set up with the standard visuals that allow you to see:

1. How much of your model is visualized
2. Where are you visualizing your model and how often is it looked at?
3. Usage based on User
4. Model Objects NOT being leveraged in Visuals
5. Objects that were added to the Report but not in your model (or old Objects that are breaking a visual)
6. All Measures & their related DAX
7. All PowerQuery steps for each query
8. Full detail for Semantic Model
9. Full detail for Visual Object Layer
