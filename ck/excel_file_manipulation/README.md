# Excel File Manipulation Request

## Problem statement
 
There are multiple excel files named high\_10000.xls and low\_10000.xls where 
the number 10000 represents the first file with increasing numbers. I need 
to take the data from particular cells in both the high and low workbooks
and store them in a new workbook with two spreadsheets. A high spreadsheet and
a low spreadsheet that corresponds to entries from the high and low workbooks, 
respectively. Each row of the new spreadsheet will correspond to data from 
the low and high file with the same number.

## Issues encountered

The files were given to me as xls files, but the python libraries I tried to 
manipulate the data didn't work. As a workaround I used libreoffice to convert 
all of the workbooks into xlsx format and then I was able to use the openpyxl 
library to complete the project.
