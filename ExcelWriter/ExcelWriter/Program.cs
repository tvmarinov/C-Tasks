using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] names = {"Ivan", "Georgi", "Radoslav", "Mariqna", "Alexandra",
                              "Hristo", "Peter", "Qvor", "Desislava", "Diana"};
            int row = 0;
            int avgScoreRow = 0;
            int rndScore = 0;
            float score = 0;
            float avgScore = 0;

            Random rnd = new Random();

            var excelApp = new Excel.Application();
            excelApp.Visible = false; // Does not open window for "Excel.exe".
            excelApp.DisplayAlerts = false;  //Does not ask permision to override, directly does it.

            Excel.Workbook book = excelApp.Workbooks.Add();
            Excel._Worksheet ws = (Excel.Worksheet)excelApp.ActiveSheet;

            ws.Cells[1, "A"] = "Name";
            ws.Cells[1, "B"] = "Age";
            ws.Cells[1, "C"] = "Score";
            ws.Cells[1, "E"] = "Average Score"; 

            ws.Range[ws.Cells[1, "A"],ws.Cells[1,"E"]].Interior.Color = Excel.XlRgbColor.rgbLightBlue;
            ws.Range[ws.Cells[1, "A"], ws.Cells[1, "E"]].EntireRow.Font.Bold = true;

            for (int i = 0; i <= 10; i++)
            {
                score = 0;

                for (int j = 2; j<=101; j++)  // First row is reserved for headers
                {

                    row = j + 100 * i;
                    rndScore = rnd.Next(0, 100);
                    ws.Cells[row, "A"] = names[rnd.Next(0, 9)];
                    ws.Cells[row, "B"] = rnd.Next(20, 80);
                    ws.Cells[row, "C"] = rndScore;
                    if(j%2 == 1)
                    {
                        ws.Range[ws.Cells[row, "A"], ws.Cells[row, "C"]].EntireRow.Font.Color = Excel.XlRgbColor.rgbLightGreen;
                    }

                    score += rndScore;
                }

                
                avgScore = score / 100;
                avgScoreRow = 2 + (100 * i);
                ws.Cells[avgScoreRow,"E"] = avgScore;
            }

            ws.Range[ws.Cells[1, "A"], ws.Cells[1, "E"]].EntireColumn.AutoFit();

            book.SaveAs("c:\\temp\\scores");
            book.Close();

            //Releasing object incase background excel processes do not close.
            Marshal.ReleaseComObject(ws); 
            Marshal.ReleaseComObject(book);


            excelApp.Quit();

        }
    }
}
