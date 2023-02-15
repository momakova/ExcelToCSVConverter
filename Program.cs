using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        Excel e1;
        // entering excel from a link given in task
        // site: https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl
        // text: Worldwide Rig Counts - Current & Historical Data
        // file name: Worldwide Rig Count Jan 2023.xlsx
        e1 = new Excel(@"https://bakerhughesrigcount.gcs-web.com/static-files/c9f87ed2-0901-4b9e-94c1-bbb874c066c0", 1);

        // for saving purpose
        e1.CreateNewFile();

        // enter starting and ending rows and starting and ending columns
        e1.ReadCell(7, 35, 2, 11);

        // possibly identifying these numbers through Latin America / Total World / Avg.
        // possibly add counter regarding years

        //saving csv file
        e1.SaveSheet("outputExcel");
    }
}
