# VBA-Challenge
This is VBA-Challenge 'The VBA of Wall Street' activity

### The VBA of Wall Street utilizes VBA scripting to analyze real stock market data.
### I have attempted the challenges given.

## Tasks:
* Create a script that will loop through all the stocks for one year and output the following information.
  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume of the stock.
* You should also have conditional formatting that will highlight positive change in green and negative change in red.

## Challenges

1. Solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 
2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Files used 

* [Test Data](alphabetical_testing.xlsx) - Used while developing scripts.
* alphabetical_testing.xlsm - macro enable workbook which contains VBA script

* [Stock Data](Multiple_year_stock_data.xlsx) - Master data - Run scripts on this data to generate   the final report.
*Multiple_year_stock_data.xlsm - macro enable workbook which contains VBA script


#### I uploaded above excel files by mistake. Please kindly ignore those 4 files. They may have been partially pushed as I exited the push process.

#### Code is included in TheVBAofWallStreet-Rupali.Surve.vbs file
#### Please refer TheVBAofWallStreet-RupaliSurve.docx document to refer detail screenshots

#### VBA Script handling ,mathematical issue

* Before starting with the script, I analyzed stock data thoroughly. There are several records within spreadsheet with zero as the opening stock price. Those records also have corresponding high, low, close and volume zero as well. This potentially indicates lack of trade for that day. However,  to calculate percent change, zero opening stock price will lead to a mathematical problem. To address this issue I have included below logic, if opening price = 0 then percent change =0 

### VBA_StockAnalysis is the main sub routine to populate summary of stock data.
### TickerSummary subroutine holds actual calculations and called from VBA_StockAnalysis.
### Summarytitles, conditionalformat, arrow sub routines are used for formatting purposes.
