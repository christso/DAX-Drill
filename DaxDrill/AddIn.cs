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
using DG2NTT.DaxDrill.Helpers;
using System.Threading;
using System.Data.SqlClient;

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

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DrillThrough")]
        public static void DrillThrough()
        {
            Task.Factory.StartNew(DrillThroughThreadSafe).ContinueWith(t =>
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (t.Exception != null)
                        MsgForm.ShowMessage(t.Exception);
                });
            });
        }
                
        public static void DrillThroughThreadSafe()
        {
            Excel.Worksheet sheet = null;
            Excel.Sheets sheets = null;
            Excel.Range rngOut = null;
            Excel.Range rngCell = null;
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            try
            {
                rngCell = excelApp.ActiveCell;

                // create sheet
                sheets = excelApp.Sheets;
                sheet = (Excel.Worksheet)sheets.Add();
                rngOut = sheet.Range["A1"];

                // set up connection
                var connString = ExcelHelper.GetConnectionString(rngCell);
                var commandText = ExcelHelper.GetDAXQuery(connString, rngCell);
                var client = new DaxClient();
                var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
                var cnn = new ADOMD.AdomdConnection(cnnStringBuilder.StrippedConnectionString);
                var dtResult = client.ExecuteTable(commandText, cnn);

                // output result to sheet
                ExcelHelper.FillRange(dtResult, rngOut);
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rngOut != null) Marshal.ReleaseComObject(rngOut);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "Get DAX Command")]
        public static void GetDrillThroughCommand()
        {
            Excel.Range rngCell = null;
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;

            try
            {
                rngCell = excelApp.ActiveCell;
                var commandText = ExcelHelper.GetDAXQuery(rngCell);
                MsgForm.ShowMessage("DAX Command", commandText);
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
            finally
            {
                if (rngCell != null) Marshal.ReleaseComObject(rngCell);
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
                string xmlString1 = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<columns xmlns=\"" + Constants.DaxDrillXmlSchemaSpace + "\">" +
                  "<column>" +
                  "<name>Call Type</name>" +
                  "<expression>Usage[Call Type]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Call Type Description</name>" +
                  "<expression>Usage[Call Type Description]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Gross Billed</name>" +
                  "<expression>Usage[Gross Billed]</expression>" +
                  "</column>" +
                  "</columns>";
                ExcelHelper.AddCustomXmlPartToWorkbook(workbook, Constants.DaxDrillXmlSchemaSpace, xmlString1);

                string xmlString2 = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<columns xmlns=\"" + Constants.DaxDrillXmlSchemaSpace + "2" + "\">" +
                  "<column>" +
                  "<name>Call Type</name>" +
                  "<expression>Usage[Call Type]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Call Type Description</name>" +
                  "<expression>Usage[Call Type Description]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Gross Billed</name>" +
                  "<expression>Usage[Gross Billed]</expression>" +
                  "</column>" +
                  "</columns>";
                ExcelHelper.AddCustomXmlPartToWorkbook(workbook, Constants.DaxDrillXmlSchemaSpace, xmlString2);


                string xmlString3 = "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                  "<columns xmlns=\"" + Constants.DaxDrillXmlSchemaSpace + "2" + "\">" +
                  "<column>" +
                  "<name>Call Type2</name>" +
                  "<expression>Usage[Call Type]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Call Type Description2</name>" +
                  "<expression>Usage[Call Type Description]</expression>" +
                  "</column>" +
                  "<column>" +
                  "<name>Gross Billed2</name>" +
                  "<expression>Usage[Gross Billed]</expression>" +
                  "</column>" +
                  "</columns>";
                ExcelHelper.AddCustomXmlPartToWorkbook(workbook, Constants.DaxDrillXmlSchemaSpace, xmlString3);
            }
            catch (Exception ex)
            {
                Helpers.MsgForm.ShowMessage(ex);
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
                string xml = ExcelHelper.ReadCustomXml(workbook, Constants.DaxDrillXmlSchemaSpace, "/x:columns");
                var columns = DaxDrillConfig.GetColumns(xml, Constants.DaxDrillXmlSchemaSpace);
            }
            catch (Exception ex)
            {
                Helpers.MsgForm.ShowMessage(ex);
            }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "Test")]
        public static void Test()
        {
            var xlApp = new Excel.Application();
            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
            xlApp.Visible = true;
            
        }
        
        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "About")]
        public static void About()
        {
            MessageBox.Show("DAX Drill is developed by DG2NTT Pty Ltd");
        }
    }
}
