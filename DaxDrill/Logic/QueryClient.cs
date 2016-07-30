using DG2NTT.DaxDrill.DaxHelpers;
using DG2NTT.DaxDrill.ExcelHelpers;
using Microsoft.AnalysisServices.Tabular;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DG2NTT.DaxDrill.Logic
{
    public class QueryClient
    {
        private readonly Excel.Range rngCell;
        public QueryClient(Excel.Range rngCell)
        {
            this.rngCell = rngCell;
        }

        public string GetDAXQuery()
        {
            var connString = ExcelHelper.GetConnectionString(rngCell);
            return GetDAXQuery(connString);
        }

        public string GetDAXQuery(string connString)
        {
            var pivotCellDic = PivotCellHelper.GetPivotCellQuery(rngCell);

            string commandText = "";
   
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);

            int maxRecords = ExcelHelper.GetMaxDrillthroughRecords(rngCell);
            var detailColumns = QueryClient.GetCustomDetailColumns(rngCell);

            using (var tabular = new DG2NTT.DaxDrill.Tabular.TabularHelper(
                cnnStringBuilder.DataSource,
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();

                // use Table Query if it exists
                // otherwise get the Table Name from the Measure

                string tableQuery = GetCustomTableQuery(rngCell);

                if (string.IsNullOrEmpty(tableQuery))
                {
                    string measureName = GetMeasureName(rngCell);
                    commandText = DaxDrillParser.BuildQueryText(tabular,
                        pivotCellDic,
                        measureName, maxRecords, detailColumns);
                }
                else
                {
                    commandText = DaxDrillParser.BuildCustomQueryText(tabular,
                        pivotCellDic,
                        tableQuery, maxRecords, detailColumns);
                }

                tabular.Disconnect();
            }

            return commandText;
        }

        public bool IsDatabaseCompatible(string connString)
        {
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
            bool result = false;

            using (var tabular = new DG2NTT.DaxDrill.Tabular.TabularHelper(
                cnnStringBuilder.DataSource,
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();
                result = tabular.IsDatabaseCompatible;
                tabular.Disconnect();
            }
            return result;
        }

        public static IEnumerable<DetailColumn> GetCustomDetailColumns(Excel.Range rngCell)
        {
            Excel.WorkbookConnection wbcnn = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;

            wbcnn = ExcelHelper.GetWorkbookConnection(rngCell);

            sheet = (Excel.Worksheet)rngCell.Parent;
            workbook = (Excel.Workbook)sheet.Parent;

            Measure measure = QueryClient.GetMeasure(rngCell);

            string xmlString = ExcelHelper.ReadCustomXmlNode(
                workbook, Constants.DaxDrillXmlSchemaSpace,
                string.Format("{0}[@id='{1}']", Constants.TableXpath, measure.Table.Name));
            List<DetailColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(Constants.DaxDrillXmlSchemaSpace, xmlString, wbcnn.Name, measure.Table.Name);

            return columns;
        }

        public static string GetCustomTableQuery(Excel.Range rngCell)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)rngCell.Parent;
            Excel.Workbook workbook = (Excel.Workbook)sheet.Parent;

            Measure measure = QueryClient.GetMeasure(rngCell);
            Office.CustomXMLNode node = ExcelHelper.GetCustomXmlNode(workbook, Constants.DaxDrillXmlSchemaSpace,
                string.Format("{0}[@id='{1}']/x:query", Constants.TableXpath, measure.Table.Name));

            if (node != null)
                return node.Text;

            return string.Empty;
        }

        private static Measure GetMeasure(Excel.Range rngCell)
        {
            var cnnString = ExcelHelper.GetConnectionString(rngCell);
            var cnnBuilder = new TabularConnectionStringBuilder(cnnString);

            string measureName = GetMeasureName(rngCell);
            Measure measure = null;
            using (var tabular = new DG2NTT.DaxDrill.Tabular.TabularHelper(cnnBuilder.DataSource, cnnBuilder.InitialCatalog))
            {
                tabular.Connect();
                measure = tabular.GetMeasure(measureName);
                tabular.Disconnect();
            }
            return measure;
        }

        public static string GetMeasureName(Excel.Range rngCell)
        {
            Excel.PivotItem pi = null;
            pi = rngCell.PivotItem;
            string piName = pi.Name;
            return DaxDrillParser.GetMeasureFromPivotItem(piName);
        }

        public static bool IsDrillThroughEnabled(Excel.Range rngCell)
        {
            Excel.PivotCache cache = null;
            Excel.PivotTable pt = null;

            try
            {
                pt = rngCell.PivotTable; // throws error if selected cell is not pivot cel
                cache = pt.PivotCache();
                if (!cache.OLAP) return false;

                // check compatibility of Tabular database
                var queryClient = new QueryClient(rngCell);
                var connString = ExcelHelper.GetConnectionString(rngCell);
                if (!queryClient.IsDatabaseCompatible(connString)) return false;

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
