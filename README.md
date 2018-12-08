# Excel-Integrator
Easily pull CSV files into existing tables enabling fast analysis updates such as pivots

# Usage Scenario
Excel-Integrator (EI) is a semi-automated procedure to include periodically updated CSV data into an Excel workbook table that serves as a source for other analysis such as pivot tables or structured table references. Let's introduce the artifacts upon which the procedure builds, summarize operation, and then outline steps to initialize the process. 

## Procedure Artifacts

1. A data source exporting CSV files readable by Excel Power Query (aka Get & Transform in Excel 2016). 
2. A Windows environment with Excel 2010 or later. I've never tested a Mac Office version supporting VBA. 
3. An Excel workbook set up to refresh from CSV (or whatever source you're able to operate) via Power Query. 
3. This same workbook with VBA macros enabled and the Table2Table macro installed and configured. 

## Operation Summary 

1. Acquire the CSV, named and located consistently along with the Excel workbook also in a consistent folder. Power Query expects this consistency as it uses absolute paths. 
2. Open and refresh the Power Query via some action such as Refresh All. All the information in the Power Query table will be replaced though Power Query does have append capabilities. 
3. Run the Table2Table macro. 
4. Refresh pivots via Refresh All. 
5. Enjoy your updated analysis. 

## Initializing the Process 

1. Install or otherwise activate Power Query which as of Office 2016, has become Get & Transform. For Office 2010, Power Query must be downloaded and installed after which it's a new ribbon item. The effort differs slightly for Office 2013 and 2016. See References below. 
2. Download and install the VBA macro, Table2Table.bas. Copy the VBA text and paste into a new module after opening the VBA Editor in Excel via Alt-F11. You also have to enable VBA macros via the File/Options/Trust Center/Trust Center Settings/Macro Settings and select Disable all macros with notification. That way, you can chose to enable macros. See References below. 

# Alternative Methods Considered

First, Excel tables have fundamental utility for managing large-ish datasets. For example, adding new data rows by typing just below the table automatically expands the table and column formulas automatically replicate for the new rows. This formula replication means you can past updated data into say, the first 10 columns, and have custom calculated columns in the next 5. I find this useful for date segregation and classification. Another fine feature, references to the tables by pivots or structured reference generally use table names. As such, once a pivot is established, refresh is easy. Maintaining these capabilities is a requirement. 

## Excel Power Query (just)

Power Query tables don't replicate formulas for new rows. 

## Google Sheets 

References such as by pivots to Sheets' version of tables do not auto-expand. Rather every pivot references cells versus by table name. As such, every pivot reference must be manually updated. 

## VBA for CSV Import 

Why not just write some VBA to open and import a CSV? Well, Power Query is just too useful to bother duplicating. It handles simple imports just fine and more complex imports including transforms and rearrangement very well. Finding Power Query was a pleasant surprise. The remaining automation task, Table to Table copy remains and yes, hosting the data twice in the same workbook is silly but... So, we use Power Query for update and VBA for the copy. See comments on Excel Power Query (just) as to why this table copy is needed. 

## Prerequisites 

Excel 2010 or higher, rights to install addins and modify VBA security settings. 

# Install & Process Initialization 

TBD

# References 

* Microsoft [Introduction to Microsoft Power Query](https://support.office.com/en-us/article/introduction-to-microsoft-power-query-for-excel-6e92e2f4-2079-4e1f-bad5-89f6269cd605)
* [Adding Code to an Excel Workbook](https://www.contextures.com/xlvba01.html)
