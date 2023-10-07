# VBA-challenge
Module 2 Challenge - VBA Scripting. A stock analysis script for formatting massive amounts of stock ticker data across multiple years.

**NOTE on sheet naming conventions:** 

The VBA script used here assumes that the names of the sheets are the years the stock data is pulled from (e.g. 2018, 2019, 2020, etc.) as that is how the base file Multiple_Year_Stock_Data.xls (not included here due to size) was formatted. It can easily be modified for a different naming scheme. 

**Docs and Resources Referenced** 
1) Code for VBA colors from: http://dmcritchie.mvps.org/excel/colors.htm

2) Code for looping through multiple worksheets inspired from Microsoft documentation: https://support.microsoft.com/en-gb/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

3) Code for improving efficiency of executed scripts (namely, Application.ScreenUpdating = False/True) from ChatGPT

4) Code to jump to top of sheet, ActiveWindow.ScrollRow = 1, from ChatGPT

5) Code to autofit columns modified from Microsoft documentation: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

6) Code to find minimum of a range heavily modified from Stack Overflow: https://stackoverflow.com/questions/37049762/finding-minimum-and-maximum-values-from-range-of-values
