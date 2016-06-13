using DG2NTT.DaxDrill.DaxHelpers;
using DG2NTT.DaxDrill.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DG2NTT.DaxDrill.Logic
{
    public class QueryLogic
    {
        public static IEnumerable<DetailColumn> GetDetailColumns(Excel.Range rngCell)
        {
            var parser = new DaxDrillParser();

            var columns = new List<DetailColumn>();
            return columns;
        }

        public static string GetDAXQuery(Excel.Range rngCell)
        {
            var connString = ExcelHelper.GetConnectionString(rngCell);
            return GetDAXQuery(connString, rngCell);
        }

        public static string GetDAXQuery(string connString, Excel.Range rngCell)
        {
            Dictionary<string, string> excelDic = PivotCellHelper.GetPivotCellQuery(rngCell);
            var parser = new DaxDrillParser();

            string commandText = "";
            string measureName = parser.GetMeasureFromPivotItem(rngCell.PivotItem.Name);
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
            var detailColumns = QueryLogic.GetDetailColumns(rngCell);

            using (var tabular = new TabularHelper(
                cnnStringBuilder.DataSource,
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();
                commandText = parser.BuildQueryText(tabular, excelDic, measureName, detailColumns);
                tabular.Disconnect();
            }

            return commandText;
        }

    }
}
