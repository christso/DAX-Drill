using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;

namespace DG2NTT.DaxDrill
{
    public class AddIn : IExcelAddIn
    {
        public void AutoClose()
        {

        }

        public void AutoOpen()
        {
            
        }

        [ExcelCommand(MenuName = "&DaxDrill", MenuText = "DrillThrough")]
        public static void DrillThrough()
        {
            var commandText = "EVALUATE TOPN ( 10, Usage)";
            var client = new DaxClient();
            var cnn = new ADOMD.AdomdConnection("Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Roaming;Data Source=localhost;");
            var result = client.ExecuteQuery(commandText, cnn);
            
        }
    }
}
