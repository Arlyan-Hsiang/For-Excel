using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace For_Excel
{
    partial class Excel_Generate
    {
        
        static void Main(string[] args)
        {
            //set the location
            string FileStr = "C:\\C# Training\\Schedule";
            //Application
            Excel.Application Excel_App1 = new Excel.Application();
            //File
            Excel.Workbook Excel_WB1 = Excel_App1.Workbooks.Add();
            //Sheet
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_WS1 = Excel_WB1.Worksheets[1];
            Excel_WS1.Name = "random";
            //record the titles
            Excel_App1.Cells[1, 1] = "Date";
            Excel_App1.Cells[1, 2] = "Recorder";
            Excel_App1.Cells[1, 3] = "Temperature";
            //datetime setting
            DateTime startdte = new DateTime(2019, 10, 1);
            DateTime enddte = new DateTime(2020, 4, 1);
            int days = Convert.ToInt16(new TimeSpan(enddte.Ticks-startdte.Ticks).Days);
            //random recorder
            string[] recorder = { "Hebe", "Darcy" };
            //random temperature
            Random rnd = new Random();
            //double randnum = NextDouble(rnd, 1, 4.5, 1);
            for (int i=2; i <= days; i++)
            {
                for (int j = 1; j <= 3; j++)
                {
                    switch (j) 
                    {
                        case 1:
                            Excel_App1.Cells[i, j] = startdte.ToString("yyyy/MM/dd");
                            startdte = startdte.AddDays(1);
                            break;
                        case 2:
                            Excel_App1.Cells[i, j] = Convert.ToString(GetRandom(recorder));
                            break;
                        case 3:
                            Excel_App1.Cells[i, j] = Convert.ToString(NextDouble(rnd,1,4.5,1) +"°C");
                            break;

                    }
                    
                }
            }

            //save
            Excel_WB1.SaveAs(FileStr);

            //close
            Excel_WS1 = null;
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_App1.Quit();
            Excel_App1 = null;
            
        }

        public static double NextDouble(Random rnd, double minValue, double maxValue, int decimalPlace)
        {
            double randNum = rnd.NextDouble() * (maxValue - minValue) + minValue;
            return Convert.ToDouble(randNum.ToString("f" + decimalPlace));
        }

        public static string GetRandom(string[] arr)
        {
            Random rnd = new Random();
            int n = rnd.Next(arr.Length);
            return arr[n];
        }

    }
    
}
