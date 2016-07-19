using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Common;
using System.Xml;
using System.Collections;
using DG2NTT.DaxDrill.DaxHelpers;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class ExcelHelper
    {
        public static string AddInPath
        {
            get
            {
                return (string)XlCall.Excel(XlCall.xlGetName);
            }
        }
        public static Excel.Worksheet AddSheet(Excel.Worksheet sh1)
        {
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sh2 = null;
            workbook = (Excel.Workbook)sh1.Parent;
            sheets = workbook.Sheets;
            sh2 = sheets.Add();
            return sh2;
        }
        
        public static string GetConnectionString(Excel.Range rngCell)
        {
            Excel.PivotTable pt = null;
            Excel.PivotCache cache = null;
            string connString;
            pt = rngCell.PivotTable;
            cache = pt.PivotCache();
            connString = cache.Connection;
            return connString;
        }

        public static Excel.WorkbookConnection GetWorkbookConnection(Excel.Range rngCell)
        {
            Excel.PivotTable pt = null;
            Excel.PivotCache cache = null;
            Excel.WorkbookConnection wbcnn = null;
            pt = rngCell.PivotTable;
            cache = pt.PivotCache();
            wbcnn = cache.WorkbookConnection;
            return wbcnn;
        }

        public static int GetMaxDrillthroughRecords(Excel.Range rngCell)
        {
            Excel.WorkbookConnection wbcnn = ExcelHelper.GetWorkbookConnection(rngCell);
            Excel.OLEDBConnection oledbcnn = wbcnn.OLEDBConnection;
            return oledbcnn.MaxDrillthroughRecords;
        }

        public static List<string> ReadCustomXmlParts(Excel.Workbook workbook)
        {
            var result = new List<string>();

            IEnumerator e = workbook.CustomXMLParts.GetEnumerator();
            Office.CustomXMLPart p;
            while (e.MoveNext())
            {
                p = (Office.CustomXMLPart)e.Current;
                //p.BuiltIn will be true for internal buildin excel parts 
                if (p != null && !p.BuiltIn)
                    result.Add(p.XML);
            }
            return result;
        }

        public static string ReadCustomXmlNode(Excel.Workbook workbook, string xNameSpace,
            string xPath)
        {
            Office.CustomXMLNode node = GetCustomXmlNode(workbook, xNameSpace, xPath);
            if (node != null)
                return node.XML;
            return string.Empty;
        }

        public static Office.CustomXMLNode GetCustomXmlNode(Excel.Workbook workbook, string xNameSpace,
            string xPath)
        {
            Office.CustomXMLParts ps = workbook.CustomXMLParts;
            ps = ps.SelectByNamespace(xNameSpace);


            for (int i = 1; i <= ps.Count; i++)
            {
                Office.CustomXMLPart p = ps[i];
                var nsmgr = p.NamespaceManager;
                nsmgr.AddNamespace("x", xNameSpace);

                Office.CustomXMLNode node = p.SelectSingleNode(xPath);
                if (node != null)
                    return node;
            }
            return null;
        }

        /*
"<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<employees xmlns=\"http://schemas.microsoft.com/vsto/samplestest\">" +
                  "<employee>" +
                  "<name>Surender GGG</name>" +
                  "<hireDate>1999-04-01</hireDate>" +
                  "<title>Manager</title>" +
                  "</employee>" +
                  "</employees>"
        */
        public static void UpdateCustomXmlPart(Excel.Workbook workbook, string namespaceName, string xmlString)
        {
            DeleteCustomXmlPart(workbook, namespaceName);
            AddCustomXmlPart(workbook, namespaceName, xmlString);
        }

        public static void UpdateCustomXmlNode(Excel.Workbook workbook, string namespaceName, string xmlString, string xPath)
        {
            Office.CustomXMLParts ps = workbook.CustomXMLParts.SelectByNamespace(namespaceName);

            foreach (Office.CustomXMLPart p in ps)
            {
                var nsmgr = p.NamespaceManager;
                nsmgr.AddNamespace("x", namespaceName);

                Office.CustomXMLNode oldNode = p.SelectSingleNode(xPath);
                oldNode.ParentNode.ReplaceChildSubtree(xmlString, oldNode);
            }
        }
        
        public static void DeleteCustomXmlPart(Excel.Workbook workbook, string namespaceName)
        {
            IEnumerator e = workbook.CustomXMLParts.GetEnumerator();
            Office.CustomXMLPart p;
            while (e.MoveNext())
            {
                p = (Office.CustomXMLPart)e.Current;
                //p.BuiltIn will be true for internal buildin excel parts 
                if (p != null && !p.BuiltIn && p.NamespaceURI == namespaceName)
                    p.Delete();
            }
        }
        public static void AddCustomXmlPart(Excel.Workbook workbook, string namespaceName, string xmlString)
        {
            System.Collections.IEnumerator enumerator = workbook.CustomXMLParts.SelectByNamespace(namespaceName).GetEnumerator();
            enumerator.Reset();

            if (!(enumerator.MoveNext()))
            {
                Office.CustomXMLPart p = workbook.CustomXMLParts.Add(xmlString);
            }
        }

        public static void FormatRange(System.Data.DataTable dataTable, Excel.Range rngOutput, int headerFlag = 1)
        {
            const int xlBaseIndex = 1;

            // get data range
            int colCnt = dataTable.Columns.Count;
            int rowCnt = dataTable.Rows.Count;
            Excel.Range rng = rngOutput.Resize[rowCnt + headerFlag, colCnt];

            // format columns
            foreach (System.Data.DataColumn column in dataTable.Columns)
            {
                // format date
                if (column.DataType == typeof(DateTime))
                {
                    Excel.Range rngColumn = rng.Columns[column.Ordinal + xlBaseIndex];
                    rngColumn.NumberFormat = "dd-mmm-yy";
                }
            }
        }

        public static void FillRange(System.Data.DataTable dataTable, Excel.Range rngOutput)
        {
            const int boundToSizeFactor = 1;
            const int rowBoundIndex = 0;
            const int columnBoundIndex = 1;

            object[,] arr = Utils.CreateArray(dataTable);
            Excel.Range rng = rngOutput.Resize[arr.GetUpperBound(rowBoundIndex) + boundToSizeFactor,
                arr.GetUpperBound(columnBoundIndex) + boundToSizeFactor];
            rng.Value2 = arr;
        }

        public static List<string> ListWorkbooks(Excel.Application excelApp)
        {
            Excel.Workbooks workbooks = excelApp.Workbooks;
            var wbList = new List<string>();
            for (int i = 1; i <= workbooks.Count; i++)
            {
                Excel.Workbook wb = workbooks[i];
                wbList.Add(wb.Name);
            }
            return wbList;
        }

        public static List<string> ListXmlNamespaces(Excel.Workbook workbook)
        {
            var result = new List<string>();
            Office.CustomXMLParts ps = workbook.CustomXMLParts;
            for (int i = 1; i <= workbook.CustomXMLParts.Count; i++)
            {
                Office.CustomXMLPart p = ps[i];

                //p.BuiltIn will be true for internal buildin excel parts 
                if (p != null && !p.BuiltIn)
                    result.Add(p.NamespaceURI);
            }

            return result;
        }

        public static Excel.Workbook FindWorkbook(string name)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbooks workbooks = excelApp.Workbooks;
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("Workbook cannot be empty");
            }
            Excel.Workbook workbook = workbooks[name];
            return workbook;
        }

        public static Excel.PivotTable CopyAsPageInvertedPivotTable(Excel.PivotTable pt)
        {
            Excel.PivotTable ptCopy = CopyPivotTable(pt);
            InvertPivotTablePageOrientation(ptCopy);
            return ptCopy;
        }

        public static Excel.PivotTable CopyPivotTable(Excel.PivotTable pt)
        {
            Excel.Application excelApp = pt.Application;
            var worksheet = (Excel.Worksheet)pt.Parent;
            worksheet.Select();
            pt.PivotSelect("", Excel.XlPTSelectionMode.xlDataAndLabel, true);
            Excel.Range rng = (Excel.Range)excelApp.Selection;
            rng.Copy();
            Excel.Worksheet ws = (Excel.Worksheet)excelApp.Sheets.Add();
            ws.Paste();
            return ws.Range["A1"].PivotTable;
        }

        public static void InvertPivotTablePageOrientation(Excel.PivotTable pt)
        {
            var cubeFields = pt.CubeFields;
            bool manualUpdate = pt.ManualUpdate;

            try
            {
                foreach (Excel.CubeField cubeField in cubeFields)
                {
                    if (cubeField.Orientation == Excel.XlPivotFieldOrientation.xlHidden
                                || cubeField.Orientation == Excel.XlPivotFieldOrientation.xlDataField)
                        continue;

                    pt.ManualUpdate = false;

                    if (cubeField.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                        cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    else
                        cubeField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                }
            }
            finally
            {
                pt.ManualUpdate = manualUpdate;
            }
        }
        
        public static bool IsMultiplePageItemsEnabled(Excel.PivotField pf)
        {

            // Excel throws error if you try to access CurrentPageName on a page field that has multiple selections
            // This function handles the error
            try
            {
                return pf.CurrentPageName == null;
            }
            catch
            {
                return pf.Orientation == Excel.XlPivotFieldOrientation.xlPageField;
            }
        }

        public static void SuppressExcelActionAlerts(Excel.Application xlApp, Action action)
        {
            bool displayAlerts = xlApp.DisplayAlerts;
            try
            {
                xlApp.DisplayAlerts = false;
                action();
            }
            finally
            {
                xlApp.DisplayAlerts = displayAlerts;
            }
        }

        public static bool IsPivotPageField(Excel.Range rngCell)
        {
            try
            {
                Excel.PivotCell pc = rngCell.PivotCell;
                Excel.PivotField pf = pc.PivotField;
                return pf.Orientation == Excel.XlPivotFieldOrientation.xlPageField;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsPivotDataCell(Excel.Range rngCell)
        {
            try
            {
                Excel.PivotCell pc = rngCell.PivotCell;
                Excel.PivotField pf = pc.PivotField;
                return pf.Orientation == Excel.XlPivotFieldOrientation.xlDataField;
            }
            catch
            {
                return false;
            }
        }

        public static void SetPivotFieldPage(Excel.PivotField pf, string currentPageName)
        {
            pf.ClearAllFilters();
            var pageName = DaxDrillParser.CreatePivotFieldPageName(pf.Name, currentPageName);
            pf.CurrentPageName = pageName;
        }
    }
}
