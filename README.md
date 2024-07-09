# VBA-Challenge
This is a multiple quarter stock Data VBA Script
Overview-This VBA script processes multiple worksheets in an Excel workbook, each containing stock data.
The script calculates and summarizes quarterly changes, percent changes, and total stock volume for each stock ticker across different quarters.
It aslo identifies and highlights the greatest percentage increase, the greatest percentage decrease, and the greatest total stock volume for each ticker.

Script Details
First we enter the Module Name
Attribute VB_Name = "Module1"

Subroutine
Sub MultipleQuarterStockData()
In this subroutine we use Dim statement to declare variables.
Dim ws As Worksheet
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim TickCount As Long
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim PerChange As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    Dim GreatIncrTicker As String
    Dim GreatDecrTicker As String
    Dim GreatVolTicker As Str
    
End Sub
Description

The MultipleQuarterStockData subroutine performs the following tasks for each worksheet in the workbook:

Initializing and Creating Column Headers:

Creates headers in columns I to L for "Ticker," "Quarterly Change," "Percent Change," and "Total Stock Volume."
Creates headers in columns P to R for identifying the tickers with the greatest percentage increase, greatest percentage decrease, and greatest total stock volume.

After that we Iterate through Rows:

The Loops iterate through each row to identify changes in quarters or tickers.
After that it calculates quarterly changes, percent changes, and total stock volume for each ticker.

Then we apply conditional formatting to highlight positive and negative changes.
Identifying the Greatest Values: Determining the tickers with the greatest percentage increase, greatest percentage decrease, and greatest total stock volume.
After that write these values and their corresponding tickers to the designated cells in in columns P to R.

Auto-fit Columns:
Adjusts column widths to fit the content.
Detailed Steps
Creating Column Headers:
Headers are added to columns I to L for ticker data and columns P to R for identifying the greatest values.
Processing Rows:
The script checks for changes in quarters or tickers, calculates quarterly changes, and writes the results to the appropriate cells.
Percent change is calculated only if the starting value is not zero to avoid division errors.
Conditional Formatting:
Cells with negative changes are colored red, positive changes are colored green, and cells with zero change have no color.
Calculating Totals and Greatest Values:
Total stock volume is summed for each quarter.
The script identifies and writes the tickers with the greatest percentage increase, decrease, and total volume.

Resizing the rows and columns for better readability.

After that we open the VBA Editor:
Press Alt + F11 to open the VBA editor in Excel.
Then right click on module and Insert the Module.
Then after we insert a new module and copy the provided script into the module.
After that we run the Script or just
Press F5 to run the script. It will process all worksheets in the workbook.
