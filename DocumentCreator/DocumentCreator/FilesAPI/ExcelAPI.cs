using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace DocumentCreator.FilesAPI
{
    public class ExcelAPI
    {
        public static Excel.Workbook GetWorkbook(string fileName)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = null;
            try
            {
                wb = app.Workbooks.Open(fileName);
            }
            catch (Exception e)
            {
                throw new Exception("Can't open file", e);
            }
            return wb;
        }

        public static Excel.Worksheet GetWorksheet(Excel.Workbook wb)
        {
            try
            {
                return wb.ActiveSheet;
            }
            catch (Exception e)
            {
                throw new Exception("Can't open worksheet", e);
            }
        }

        public static Excel.Worksheet GetWorksheet(Excel.Workbook wb, int worksheet)
        {
            try
            {
                return wb.Sheets[worksheet];
            }
            catch(Exception e)
            {
                throw new Exception("Not correct number sheet", e);
            }
        }

        public static Excel.Worksheet GetWorksheet(Excel.Workbook wb, string worksheet)
        {
            try
            {
                return wb.Sheets[worksheet];
            }
            catch (Exception e)
            {
                throw new Exception("Not correct name of sheet", e);
            }
        }

        public static void saveFile(Excel.Workbook wb, string fileName = "")
        {
            if (string.IsNullOrEmpty(fileName))
            {
                try
                {
                    wb.Save();
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file", e);
                }
            }
            else
            {
                try
                {
                    wb.SaveAs(fileName);
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file in "+fileName, e);
                }
            }
        }

        public static void close(Excel.Workbook wb)
        {
            if (wb != null)
            {
                wb.Close();
                killExcel();
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void killExcel()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in procs)
            {
                p.Kill();
            }
        }
    }
}
