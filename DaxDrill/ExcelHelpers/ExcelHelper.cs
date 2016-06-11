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
using DG2NTT.DaxDrill.Helpers;
using System.Data.SqlClient;
using System.Data.Common;
using System.Xml;
using System.Collections;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class ExcelHelper
    {
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
            }
            finally
            {
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (cache != null) Marshal.ReleaseComObject(cache);
            }
            return connString;
        }
        public static string GetDAXQuery(Excel.Range rngCell)
        {
            var connString = GetConnectionString(rngCell);
            return GetDAXQuery(connString, rngCell);
        }

        public static string FormatConnectionString(string connString)
        {
            var cnnStringBuilder = new DbConnectionStringBuilder();
            cnnStringBuilder.ConnectionString = connString;
            string dataSource = cnnStringBuilder["data source"].ToString();
            string initialCatalog = cnnStringBuilder["initial catalog"].ToString();
            return string.Format(
                "Integrated Security=SSPI;Persist Security Info=True;Initial Catalog={1};Data Source={0};", 
                dataSource, initialCatalog);
        }

        public static string GetDAXQuery(string connString, Excel.Range rngCell)
        {
            Dictionary<string, string> excelDic = PivotCellHelper.GetPivotCellQuery(rngCell);
            var parser = new DaxDrillParser();

            string commandText = "";
            string measureName = parser.GetMeasureFromPivotItem(rngCell.PivotItem.Name);
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);

            using (var tabular = new TabularHelper(
                cnnStringBuilder.DataSource, 
                cnnStringBuilder.InitialCatalog))
            {
                tabular.Connect();
                commandText = parser.BuildQueryText(tabular, excelDic, measureName);
                tabular.Disconnect();
            }

            return commandText;
        }

        public static string ReadCustomXmlPart(Excel.Workbook workbook, string xNameSpace,
            string xPath)
        {
            System.Collections.IEnumerator enumerator = workbook.CustomXMLParts.SelectByNamespace(Constants.DaxDrillXmlSchemaSpace).GetEnumerator();
            enumerator.Reset();
            while (enumerator.MoveNext())
            {
                Office.CustomXMLPart p = (Office.CustomXMLPart)enumerator.Current;
                p.NamespaceManager.AddNamespace("x", xNameSpace);
                Office.CustomXMLNode node = p.SelectSingleNode(xPath);
                return node.XML;
            }
            return string.Empty;
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
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rng != null) Marshal.ReleaseComObject(rng);
            }
        }


    }
}
