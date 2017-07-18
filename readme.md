## vb-read-write-sheet

This is a minimal VB.net Smartsheet sample that demonstrates how to
* Load a sheet
* Loop through the rows
* Check for rows that meet a criteria
* Update cell values
* Write the results back to the original sheet


This sample scans a sheet for rows where the value of the "Status" column is "Complete" and sets the "Remaining" column to zero.
This is implemented in the `evaluateRowAndBuildUpdates()` method which you should modify to meet your needs.


## Setup
Import the sample data from "Sample Sheet.xlsx" into a new sheet

Update the App.config file with these two settings:
* An API access token, obtained from the Smartsheet Account button, under Personal settings
* The Sheet Id, obtained from sheet properties 

Build and run the application.

The rows marked "Complete" will have the "Remaining" value set to 0. (Note that you will have to refresh in the desktop application to see the changes)


## Details
Since all .net languages are interoperable, we can use the C# SDK from Visual Basic.
To add a reference to the SDK, right-click the project in Solution Explorer and pick Manage NuGet Packages. Browse for 'smartsheet' to add a reference to the `smartsheet-csharp-sdk`.

Also, add the following lines to your source code:

    Imports Smartsheet.Api
    Imports Smartsheet.Api.Models

## See also
- https://github.com/smartsheet-platform/smartsheet-csharp-sdk
- http://smartsheet-platform.github.io/api-docs/
- https://www.smartsheet.com/
