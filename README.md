This is an old repo that still works but a much more comprehensive solution is here: https://github.com/chris1642/Power-BI-Backup-Impact-Analysis-Governance-Solution

Recommended: use a Base folder that is cloud-based (such as a SharePoint folder synced to your computer) - this will make it easy to leverage within a Power BI semantic model and automatic refreshes. 

1. PowerShell Script: Define Workspace ID & Base Folder (currently defaulted to C:/Power BI Backups)
2. C# Script: Update Base Folder (only necessary if you change the Base folder from the default in Step 1)
3. Place the 'Blank Model.bim' in the same Base Folder defined in Step 1 

Run PowerShell Script. This will auto install Power BI & Excel modules (if not already installed), download all your reports from the Workspace ID defined in step 1, have Tabular Editor 2 run the C# script to extract all of the visual object layer details, and then PowerShell will combine the outputs into a single Excel file. 

Once the base folder is defined, this can be re-ran as often as you'd like, and the combined Excel file will append with the associated date related to the Reports. 


Optional:

Step 1: Connect to a Semantic Model in Tabular Editor, go to the C# script window and run the Optional Step 1 Data Export. This will export all of the model details into csv files in the same base folder + Model Backups

Step 2: Any models exported will be combined into a single Excel file. 







Once the refresh is completed, the measures and relationships are set up with the standard visuals that allow you to see:

1. How much of your model are you actually using in visuals or filters?
2. Where are you visualizing your model and how often is it looked at?
3. Usage based on User, Report, and Page
4. Model Objects NOT being used in Visuals or Filters
5. Objects that were added to the Report but are not in your model (or old Objects that are breaking a visual)
6. All Measures & their related DAX
7. All PowerQuery steps for each query
8. Full detail for Semantic Model
9. Full detail for Visual Object Layer
