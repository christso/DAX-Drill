using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADOMD = Microsoft.AnalysisServices.AdomdClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DG2NTT.DaxDrill.ExcelHelpers;

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

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DrillThrough")]
        public static void DrillThrough()
        {
            Excel.Worksheet sheet = null;
            Excel.Sheets sheets = null;
            Excel.Range rngOut = null;
            Excel.Range rngCell = null;
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            try
            {
                // set up connection
                rngCell = excelApp.ActiveCell;
                var connString = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=Roaming;Data Source=localhost;";
                var commandText = ExcelHelper.GetDAXQuery(rngCell);
                var client = new DaxClient();
                var cnn = new ADOMD.AdomdConnection(connString);
                var dtResult = client.ExecuteTable(commandText, cnn);

                // output result to new sheet
                sheets = excelApp.Sheets;
                sheet = (Excel.Worksheet)sheets.Add();
                rngOut = sheet.Range["A1"];
                ExcelHelper.FillRange(dtResult, rngOut);
            }
            catch (Exception ex)
            {
                Helpers.ErrForm.ShowException(ex);
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rngOut != null) Marshal.ReleaseComObject(rngOut);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "Add XML")]
        public static void AddXML()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbook = excelApp.ActiveWorkbook;
                string xmlString = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<employees xmlns=\"http://schemas.microsoft.com/vsto/samplestest\">" +
                  "<employee>" +
                  "<name>Surender GGG</name>" +
                  "<hireDate>1999-04-01</hireDate>" +
                  "<title>Manager</title>" +
                  "</employee>" +
                  "</employees>";
                ExcelHelper.AddCustomXmlPartToWorkbook(workbook, xmlString);
            }
            catch (Exception ex)
            {
                Helpers.ErrForm.ShowException(ex);
            }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "Read XML")]
        public static void ReadXML()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbook = excelApp.ActiveWorkbook;
                ExcelHelper.ReadCustomXmlPartSingleNode(workbook);
            }
            catch (Exception ex)
            {
                Helpers.ErrForm.ShowException(ex);
            }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "About")]
        public static void About()
        {
            MessageBox.Show("DAX Drill is developed by DG2NTT Pty Ltd");
        }
    }
}
