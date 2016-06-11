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

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class ExcelHelper
    {
        public static string GetDAXQuery(Excel.Range rngCell)
        {
            Dictionary<string, string> excelDic = PivotCellHelper.GetPivotCellQuery(rngCell);
            var parser = new DaxDrillParser();

            string commandText = "";
            string measureName = parser.GetMeasureFromPivotItem(rngCell.PivotItem.Name);

            // get connection from Pivot Table
            using (var tabular = new TabularHelper("localhost", "Roaming"))
            {

                tabular.Connect();
                commandText = parser.BuildQueryText(tabular, excelDic, measureName);
                tabular.Disconnect();
            }

            #region DEBUG
            foreach (var pair in excelDic)
            {
                Debug.Print("{0} | {1}", pair.Key, pair.Value);
            }
            #endregion

            return commandText;
        }

        public static void ReadCustomXmlPartSingleNode(Excel.Workbook workbook)
        {
            System.Collections.IEnumerator enumerator = workbook.CustomXMLParts.SelectByNamespace("http://schemas.microsoft.com/vsto/samplestest").GetEnumerator();
            enumerator.Reset();
            if (enumerator.MoveNext())
            {
                Office.CustomXMLPart a = (Office.CustomXMLPart)enumerator.Current;
                a.NamespaceManager.AddNamespace("x", "http://schemas.microsoft.com/vsto/samplestest");
                MessageBox.Show(a.SelectSingleNode("/x:employees/x:employee/x:name").Text);
                MessageBox.Show(a.XML);
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
        public static void AddCustomXmlPartToWorkbook(Excel.Workbook workbook, string xmlString)
        {

            System.Collections.IEnumerator enumerator = workbook.CustomXMLParts.SelectByNamespace("http://schemas.microsoft.com/vsto/samplestest").GetEnumerator();
            enumerator.Reset();

            if (!(enumerator.MoveNext()))
            {
                string xmlString1 = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<employees xmlns=\"http://schemas.microsoft.com/vsto/samplestest\">" +
                  "<employee>" +
                  "<name>Surender GGG</name>" +
                  "<hireDate>1999-04-01</hireDate>" +
                  "<title>Manager</title>" +
                  "</employee>" +
                  "</employees>";

                Office.CustomXMLPart employeeXMLPart = workbook.CustomXMLParts.Add(xmlString1);

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
