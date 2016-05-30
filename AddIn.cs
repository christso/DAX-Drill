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
            Excel.Range rng = null;
            Excel.Application app = null;

            try
            {
                var connString = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Roaming;Data Source=localhost;";
                var commandText = "EVALUATE TOPN ( 10, Usage)";
                var client = new DaxClient();
                var cnn = new ADOMD.AdomdConnection(connString);
                var dtResult = client.ExecuteQuery(commandText, cnn);

                Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var excelHelper = new ExcelHelper(excelApp);

                sheet = (Excel.Worksheet)excelApp.ActiveSheet;
                rng = sheet.Range["A1"];
                excelHelper.FillRange(dtResult, rng);
            }
            finally
            {
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rng != null) Marshal.ReleaseComObject(rng);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }


        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "About")]
        public static void About()
        {
            MessageBox.Show("DAX Drill is developed by DG2NTT Pty Ltd");
        }
    }
}
