using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace RTUHM.Utility
{
    class ExcelBuilder
    {
        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
        private List<Microsoft.Office.Interop.Excel.Worksheet> xlWorksheets = new List<Microsoft.Office.Interop.Excel.Worksheet>();

        private string fullFilePath = String.Empty;


        private List<int> worksheetRowCount = new List<int>();

        public ExcelBuilder(string filename) 
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                throw new Exception("Excel is not installed properly on this machine");
            } 
            else 
            {
                fullFilePath = filename + string.Format("-{0:yyyy-MM-dd_hh-mm-ss-tt}.xlsx", DateTime.Now);
            }
        }

        public void OpenWorkbook()
        {
            xlWorkbook = xlApp.Workbooks.Add();
        }

        public int CreateNewWorksheet(string worksheetName = null)
        {
            int worksheetID = 0;

            xlWorksheets.Add(xlWorkbook.Worksheets.Add());
            worksheetID = xlWorksheets.Count - 1;

            if (worksheetName != null)
            {
                xlWorksheets[worksheetID].Name = worksheetName;
            }

            worksheetRowCount.Add(1);
            return worksheetID;
        }

        public void WriteHeaderToWorksheet(int worksheetID, string[] headers)
        {
            for (int column = 0; column < headers.Length; column++)
            {
                xlWorksheets[worksheetID].Cells[worksheetRowCount[worksheetID], column + 1] = headers[column];
                xlWorksheets[worksheetID].Cells[worksheetRowCount[worksheetID], column + 1].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGray;
                xlWorksheets[worksheetID].Cells[worksheetRowCount[worksheetID], column + 1].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;
            }

            worksheetRowCount[worksheetID]++;
        }

        public void WriteDataToWorksheet(int worksheetID, string[] data)
        {
            for (int column = 0; column < data.Length; column++)
            {
                xlWorksheets[worksheetID].Cells[worksheetRowCount[worksheetID], column + 1] = data[column];
            }
            worksheetRowCount[worksheetID]++;
        }

        public void CloseWorkbook() 
        {
            xlWorkbook.SaveAs(Path.GetFullPath(fullFilePath));
            xlWorkbook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();

            foreach (Microsoft.Office.Interop.Excel.Worksheet xlWorksheet in xlWorksheets)
            {
                CleanupCOMObject(xlWorksheet);
            }
            CleanupCOMObject(xlWorkbook);
            CleanupCOMObject(xlApp);
        }

        private void CleanupCOMObject(object obj) 
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        public string GetFilename()
        {
            return fullFilePath;
        }
    }
}
