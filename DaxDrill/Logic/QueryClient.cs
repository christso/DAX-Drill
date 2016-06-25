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
            Dictionary<string, string> excelDic = PivotCellHelper.GetPivotCellQuery(rngCell);

            string commandText = "";
            string measureName = GetMeasureName(rngCell);
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);

            int maxRecords = ExcelHelper.GetMaxDrillthroughRecords(rngCell);
            var detailColumns = QueryClient.GetDetailColumns(rngCell);

            using (var tabular = new TabularHelper(
                cnnStringBuilder.DataSource,
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();
                commandText = DaxDrillParser.BuildQueryText(tabular, excelDic, measureName, maxRecords, detailColumns);
                tabular.Disconnect();
            }

            return commandText;
        }

        public static IEnumerable<DetailColumn> GetDetailColumns(Excel.Range rngCell)
        {

            Excel.WorkbookConnection wbcnn = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
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
            finally
            {
                if (wbcnn != null) Marshal.ReleaseComObject(wbcnn);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        public static IEnumerable<DetailColumn> GetDetailColumns2(Excel.Range rngCell)
        {

            Excel.WorkbookConnection wbcnn = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                
                wbcnn = ExcelHelper.GetWorkbookConnection(rngCell);

                sheet = (Excel.Worksheet)rngCell.Parent;
                workbook = (Excel.Workbook)sheet.Parent;

                var cnnString = ExcelHelper.GetConnectionString(rngCell);
                var cnnBuilder = new TabularConnectionStringBuilder(cnnString);

                /*string measureName = GetMeasureName(rngCell); */
                /*
                Measure measure = null;
                using (var tabular = new TabularHelper(cnnBuilder.DataSource, cnnBuilder.InitialCatalog))
                {
                    tabular.Connect();
                    measure = tabular.GetMeasure(measureName);
                    tabular.Disconnect();
                }
                */
                return null;
            }
            finally
            {
                if (wbcnn != null) Marshal.ReleaseComObject(wbcnn);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
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
            try
            {
                pi = rngCell.PivotItem;
                string piName = pi.Name;
                return DaxDrillParser.GetMeasureFromPivotItem(piName);
            }
            finally
            {
                Marshal.ReleaseComObject(pi);
            }
        }

        public static bool IsDrillThroughEnabled(Excel.Range rngCell)
        {
            Excel.PivotCache cache = null;
            Excel.PivotTable pt = null;

            try
            {
                pt = rngCell.PivotTable;
                cache = pt.PivotCache();
                return cache.OLAP;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (cache != null) Marshal.ReleaseComObject(cache);
            }

        }

    }
}
