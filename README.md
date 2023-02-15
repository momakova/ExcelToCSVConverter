# ExcelToCSVConverter

#Abount#
This console app downloads Excel file from a link and converts from 7th to 35th row, from 2nd to 11th column (which is data for past two years) into a CSV file.
I am using Microsoft.Office.Interop.Excel library
App contains 3 separate classes - Main (Program), Excel and Downloader

#Issues#
I had major problem with downloading Excel file from site, so I added possibility to run code with link from microsoft site where xlsx file example is downloaded.
Just delete comment sign from line
url = "https://go.microsoft.com/fwlink/?LinkID=521962";
in Downloader class.

#Backup#
Added backup possibility - create CSV based on previously downloaded xlsx file from the site
