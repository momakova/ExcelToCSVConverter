using System;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

public class Excel
{
    //path and directory configuration
    string appDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
    string path = "";

    //definition of excel !!! add ref to Microsoft Excel Object Library
    _Application excel = new _Excel.Application();
    Workbook wb;
    Worksheet ws;
    Workbook wb1;
    Worksheet ws1;

    //creating excel sheet from a link
    public Excel(string link, int Sheet)
    {
        Downloader d = new Downloader();
        path = d.GetPath(link);
        wb = excel.Workbooks.Open(path);
        ws = wb.Worksheets[Sheet];
    }

    // creating alternative CSV workbook
    public void CreateNewFile()
    {
        wb1 = excel.Workbooks.Add();
        ws1 = wb1.Worksheets[1];
    }
    
    // go through excel and fill CSV
    public void ReadCell(int startR, int endR, int startC, int endC)
    {
        //defining parrameters
        int r = startR;
        int rows = endR;
        int c = startC;
        int columns = endC;
        // for CSV
        int r1 = 1;
        int c1 = 1;
        //going through table
        for (; r <= rows; r++)
        {
            while (ws.Cells[r, c].Value != null)
            {
                // writing data
                Console.Write(ws.Cells[r, c].Value.ToString() + ",");
                // adding data to csv table
                ws1.Cells[r1, c1].Value = ws.Cells[r, c].Value;
                c++;
                c1++;
            }
            for (; c <= columns; c++)
            {
                Console.Write(",");
                ws1.Cells[r1, c1].Value = null;
                c1++;
            }
            Console.Write("\n");
            c = startC;
            c1 = 1;
            r1++;
        }
    }


    // method to save csv file
    public void SaveSheet(string name)
    {
        // naming CSV file
        Random random = new Random();
        int randomNumber = random.Next(1, 999);
        string id = name + randomNumber;

        // saving csv file
        ws1.SaveAs(appDirectory + @"..\..\..\data\" + id + @".csv");
        Console.WriteLine("\nSaved to my data folder as " + id + @".csv");
        Console.WriteLine("You can find table in project directory in data folder.");
    }
}
