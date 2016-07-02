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

namespace DG2NTT.DaxDrill.Logic
{
    public class QueryClient
    {
        private readonly Excel.Range rngCell;
        public QueryClient(Excel.Range rngCell)
        {
            this.rngCell = rngCell;
 
        }

        public string GetDaxQuery()
        {
            return GetDAXQuery(this.rngCell);
        }

        public string GetDAXQuery(Excel.Range rngCell)
        {
            var connString = ExcelHelper.GetConnectionString(rngCell);
            return GetDAXQuery(connString, rngCell);
        }

        public string GetDAXQuery(string connString, Excel.Range rngCell)
        {
            var pivotCellDic = PivotCellHelper.GetPivotCellQuery(rngCell);

            string commandText = "";
            string measureName = GetMeasureName(rngCell);
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);

            int maxRecords = ExcelHelper.GetMaxDrillthroughRecords(rngCell);
            var detailColumns = QueryClient.GetCustomDetailColumns(rngCell);

            using (var tabular = new TabularHelper(
                cnnStringBuilder.DataSource,
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();
                commandText = DaxDrillParser.BuildQueryText(tabular, 
                    pivotCellDic, 
                    measureName, maxRecords, detailColumns);
                tabular.Disconnect();
            }

            return commandText;
        }

        public bool IsDatabaseCompatible(string connString)
        {
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
            bool result = false;

            using (var tabular = new TabularHelper(
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

            string xmlString = ExcelHelper.ReadCustomXmlPart(
                workbook, Constants.DaxDrillXmlSchemaSpace, 
                Constants.TableXpath);
            List<DetailColumn> columns = DaxDrillConfig.GetColumnsFromTableXml(
                            wbcnn.Name, measure.Table.Name, xmlString, Constants.DaxDrillXmlSchemaSpace);

            return columns;
        }
        
        private static Measure GetMeasure(Excel.Range rngCell)
        {
            var cnnString = ExcelHelper.GetConnectionString(rngCell);
            var cnnBuilder = new TabularConnectionStringBuilder(cnnString);

            string measureName = GetMeasureName(rngCell);
            Measure measure = null;
            using (var tabular = new TabularHelper(cnnBuilder.DataSource, cnnBuilder.InitialCatalog))
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
