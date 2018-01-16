using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace codeFestOwn
{
    class fetchData
    {
        // To fetch data from the excel worksheets
        public String[] fetchKeysFromExcel(int flag)
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlsApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                //return null;
            }

            //Displays Excel so you can see what is happening

            if (flag == 1)
            {
                Workbook wb = xlsApp.Workbooks.Open(Program.excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
                Sheets sheets = wb.Worksheets;
                Worksheet wsD1 = (Worksheet)sheets.get_Item(1);
                Worksheet wsD2 = (Worksheet)sheets.get_Item(2);
                Range firstColumnD1 = wsD1.UsedRange.Columns[1];
                Range firstColumnD2 = wsD2.UsedRange.Columns[1];
                System.Array myvaluesD1 = (System.Array)firstColumnD1.Cells.Value;
                string[] strArrayD1 = myvaluesD1.OfType<object>().Select(o => o.ToString()).ToArray();

                return strArrayD1;
            }
            else
            {
                Workbook wb = xlsApp.Workbooks.Open(Program.excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true);
                Sheets sheets = wb.Worksheets;
                Worksheet wsD1 = (Worksheet)sheets.get_Item(1);
                Worksheet wsD2 = (Worksheet)sheets.get_Item(2);
                Range firstColumnD1 = wsD1.UsedRange.Columns[1];
                Range firstColumnD2 = wsD2.UsedRange.Columns[1];
                System.Array myvaluesD2 = (System.Array)firstColumnD2.Cells.Value;
                string[] strArrayD2 = myvaluesD2.OfType<object>().Select(o => o.ToString()).ToArray();

                return strArrayD2;
            }
        }
    }
}
