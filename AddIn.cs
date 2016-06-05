using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DG2NTT.DaxDrill.ExcelHelpers;

namespace DG2NTT.DaxDrill
{
    public class AddIn : IExcelAddIn
    {
        public void AutoClose()
        {
            var x = 1;
        }

        public void AutoOpen()
        {
            
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DrillThrough")]
        public static void DrillThrough()
        {
            Excel.Worksheet sheet = null;
            Excel.Sheets sheets = null;
            Excel.Range rngOut = null;
            Excel.Range rngCell = null;
            Excel.Application app = null;

            try
            {
                // set up Excel Helper
                Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var excelHelper = new ExcelHelper(excelApp);

                // set up connection
                rngCell = excelApp.ActiveCell;
                var connString = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Roaming;Data Source=localhost;";
                var commandText = excelHelper.GetDAXQuery(rngCell);
                var client = new DaxClient();
                var cnn = new ADOMD.AdomdConnection(connString);
                var dtResult = client.ExecuteQuery(commandText, cnn);

                // output result to new sheet
                sheets = excelApp.Sheets;
                sheet = (Excel.Worksheet)sheets.Add();
                rngOut = sheet.Range["A1"];
                excelHelper.FillRange(dtResult, rngOut);
            }
            catch (Exception ex)
            {
                Helpers.ErrForm.ShowException(ex);
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rngOut != null) Marshal.ReleaseComObject(rngOut);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "Test")]
        public static void TestErrForm()
        {
            var ex = new InvalidOperationException("This is a test error");
            Helpers.ErrForm.ShowException(ex);
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "About")]
        public static void About()
        {
            MessageBox.Show("DAX Drill is developed by DG2NTT Pty Ltd");
        }
    }
}
