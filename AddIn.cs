using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using Excel = Microsoft.Office.Interop.Excel;

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

        [ExcelCommand(MenuName = "&DaxDrill", MenuText = "DrillThrough")]
        public static void DrillThrough()
        {

            var connString = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Roaming;Data Source=localhost;";
            var commandText = "EVALUATE TOPN ( 10, Usage)";
            var client = new DaxClient();
            var cnn = new ADOMD.AdomdConnection(connString);
            var dtResult = client.ExecuteQuery(commandText, cnn);

            var excelApp = (Excel.Application)ExcelDnaUtil.Application;
            var excelHelper = new ExcelHelper(excelApp);

            var sheet = (Excel.Worksheet)excelApp.ActiveSheet;
            excelHelper.FillRange(dtResult, sheet.Range["A1"]);
        }
    }
}
