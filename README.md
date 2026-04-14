# Import Data from SQL Server to Excel Table in C#

This sample demonstrates how to import data from a SQL Server database into an Excel table using C# and Syncfusion XlsIO. It shows a complete workflow-from establishing a database connection to exporting structured data into a formatted Excel table-making it useful for reporting, data analysis, and automation scenarios. In addition to the basic import process, the sample highlights how developers can work with query parameters, refresh Excel tables to pull updated values, and apply formatting for improved readability. It also illustrates how constant, range, and prompt parameters can be used to filter data dynamically, enabling flexible reporting directly within Excel. By leveraging Syncfusion’s powerful XlsIO library, this project eliminates the need for manual data transfers, supports automation of recurring tasks, and ensures that business users can generate professional Excel reports with minimal effort.

## What this sample covers
* Connects to SQL Server using ADO.NET
* Executes SQL queries to retrieve data
* Loads data into an Excel worksheet
* Converts the data range into an Excel Table
* Saves the Excel file without requiring Microsoft Excel

## Prerequisites
* .NET Framework / .NET
* SQL Server with sample data
* Syncfusion XlsIO NuGet package

## Use cases
* Exporting reports from databases
* Automating Excel report generation
* Creating structured Excel tables for business analysis

## Repository Structure
* **Query-Parameters:** Demonstrates how to use constant, range, and prompt parameters in SQL queries bound to Excel tables. This allows developers to filter data dynamically and interactively, making reports more flexible.
* **SQL-Server-To-Excel-Table:** Shows how to establish a direct connection to SQL Server, execute queries, and refresh Excel tables with updated values. This ensures that Excel reports stay synchronized with the underlying database.

## Key Highlights
* **Constant Parameters:** Pass fixed values into SQL queries for repeatable filtering.
* **Range Parameters:** Bind query parameters to Excel cell ranges, enabling dynamic filtering based on user input.
* **Prompt Parameters:** Allow interactive input at runtime, making the solution adaptable to different reporting needs.
* **Refresh Support:** Excel tables can be refreshed to pull the latest data from SQL Server without recreating the workbook.
* **Formatting:** Columns are auto-fitted, and tables are styled to improve readability.

This project serves as a simple, production-ready reference for developers looking to integrate SQL Server data with Excel using C#.