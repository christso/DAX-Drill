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
        public static Excel.Worksheet AddSheet(Excel.Worksheet sh1)
        {
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sh2 = null;
            try
            {
                workbook = (Excel.Workbook)sh1.Parent;
                sheets = workbook.Sheets;
                sh2 = sheets.Add();
                return sh2;
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (sh2 != null) Marshal.ReleaseComObject(sh2);
            }
            
        }
        
        public static string GetConnectionString(Excel.Range rngCell)
        {
            Excel.PivotTable pt = null;
            Excel.PivotCache cache = null;
            string connString;
            try
            {
                pt = rngCell.PivotTable;
                cache = pt.PivotCache();
                connString = cache.Connection;
                return connString;
            }
            finally
            {
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (cache != null) Marshal.ReleaseComObject(cache);
            }
        }

        public static Excel.WorkbookConnection GetWorkbookConnection(Excel.Range rngCell)
        {
            Excel.PivotTable pt = null;
            Excel.PivotCache cache = null;
            Excel.WorkbookConnection wbcnn = null;
            try
            {
                pt = rngCell.PivotTable;
                cache = pt.PivotCache();
                wbcnn = cache.WorkbookConnection;
                return wbcnn;
            }
            finally
            {
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (cache != null) Marshal.ReleaseComObject(cache);
            }
        }

        public static int GetMaxDrillthroughRecords(Excel.Range rngCell)
        {
            Excel.WorkbookConnection wbcnn = null;
            Excel.OLEDBConnection oledbcnn = null;
            try
            {
                wbcnn = ExcelHelper.GetWorkbookConnection(rngCell);
                oledbcnn = wbcnn.OLEDBConnection;
                return oledbcnn.MaxDrillthroughRecords;
            }
            finally
            {
                if (wbcnn != null) Marshal.ReleaseComObject(wbcnn);
                if (oledbcnn == null) Marshal.ReleaseComObject(oledbcnn);
            }
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

                Marshal.ReleaseComObject(p);
            }
                
            return result;
        }

        public static string ReadCustomXmlPart(Excel.Workbook workbook, string xNameSpace,
            string xPath)
        {
            Office.CustomXMLParts ps = null;

            try
            {
                ps = workbook.CustomXMLParts;
                ps = ps.SelectByNamespace(xNameSpace);

                for (int i = 1; i <= ps.Count; i++)
                {
                    Office.CustomXMLPart p = ps[i];
                    var nsmgr = p.NamespaceManager;
                    nsmgr.AddNamespace("x", xNameSpace);
                    Office.CustomXMLNode node = p.SelectSingleNode(xPath);

                    Marshal.ReleaseComObject(nsmgr);
                    Marshal.ReleaseComObject(p);

                    if (node != null)
                    {
                        var xml = node.XML;
                        Marshal.ReleaseComObject(node);
                        return xml;
                    }
                }
                return string.Empty;
            }
            finally
            {
                if (ps != null) Marshal.ReleaseComObject(ps);
            }
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

        public static void FillRange(System.Data.DataTable dataTable, Excel.Range rngOutput)
        {
            Excel.Application excelApp = rngOutput.Application;
            Excel.Worksheet sheet = excelApp.ActiveSheet;
            Excel.Range rng = null;
            const int boundToSizeFactor = 1;
            const int rowBoundIndex = 0;
            const int columnBoundIndex = 1;

            try
            {

                object[,] arr = Utils.CreateArray(dataTable);
                rng = rngOutput.Resize[arr.GetUpperBound(rowBoundIndex) + boundToSizeFactor,
                    arr.GetUpperBound(columnBoundIndex) + boundToSizeFactor];
                rng.Value2 = arr;

            }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rng != null) Marshal.ReleaseComObject(rng);
            }
        }

        public static List<string> ListWorkbooks(Excel.Application excelApp)
        {
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = excelApp.Workbooks;
                var wbList = new List<string>();
                for (int i = 1; i <= workbooks.Count; i++)
                {
                    Excel.Workbook wb = workbooks[i];
                    wbList.Add(wb.Name);
                    Marshal.ReleaseComObject(wb);
                }
                return wbList;
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            }
        }

        public static List<string> ListXmlNamespaces(Excel.Workbook workbook)
        {
            Office.CustomXMLParts ps = null;
            try
            {
                var result = new List<string>();
                ps = workbook.CustomXMLParts;
                for (int i = 1; i <= workbook.CustomXMLParts.Count; i++)
                {
                    Office.CustomXMLPart p = ps[i];

                    //p.BuiltIn will be true for internal buildin excel parts 
                    if (p != null && !p.BuiltIn)
                        result.Add(p.NamespaceURI);

                    Marshal.ReleaseComObject(p);
                }

                return result;
            }
            finally
            {
                if (ps != null) Marshal.ReleaseComObject(ps);
            }
        }

        public static Excel.Workbook FindWorkbook(string name)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Workbooks workbooks = null;
            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbooks = excelApp.Workbooks;
                if (string.IsNullOrWhiteSpace(name))
                {
                    throw new InvalidOperationException("Workbook cannot be empty");
                }
                workbook = workbooks[name];
                return workbook;
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

    }
}
