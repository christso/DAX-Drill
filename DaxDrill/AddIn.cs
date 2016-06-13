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
using System.Threading;
using System.Data.SqlClient;
using System.Diagnostics;
using DG2NTT.DaxDrill.UI;
using DG2NTT.DaxDrill.Logic;
using DG2NTT.DaxDrill.DaxHelpers;

namespace DG2NTT.DaxDrill
{
    public class AddIn : IExcelAddIn
    {
        // set to true to force Excel to close
        public const bool KillExcel = false;
        public void AutoClose()
        {
            var excelApp = (Excel.Application)ExcelDnaUtil.Application;
            excelApp.WorkbookDeactivate -= XlApp_WorkbookDeactivate;
        }

        public void AutoOpen()
        {
            var excelApp = (Excel.Application)ExcelDnaUtil.Application;
            excelApp.WorkbookDeactivate += XlApp_WorkbookDeactivate;
        }

        // kill Excel process in case objects are not properly released
        private void XlApp_WorkbookDeactivate(Excel.Workbook Wb)
        {
            if (KillExcel && Wb.Application.Workbooks.Count == 1)
            {
                Process.GetCurrentProcess().Kill();
            }
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
                var commandText = QueryLogic.GetDAXQuery(connString, rngCell);
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
                if (rngCell != null) Marshal.ReleaseComObject(rngCell);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DAX Query")]
        public static void DrillThroughQuery()
        {
            Excel.Range rngCell = null;
            Excel.Application excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook workbook = null;

            try
            {
                // XML configuration
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbook = excelApp.ActiveWorkbook;
                string xml = ExcelHelper.ReadCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, "/x:columns");
                var columns = DaxHelpers.DaxDrillConfig.GetColumnsFromColumnsXml(xml, Constants.DaxDrillXmlSchemaSpace);

                // generate command
                rngCell = excelApp.ActiveCell;
                var commandText = QueryLogic.GetDAXQuery(rngCell);
                MsgForm.ShowMessage("DAX Query", commandText);
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
            finally
            {
                if (rngCell != null) Marshal.ReleaseComObject(rngCell);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "XML Metadata")]
        public static void ShowMetadataEditor()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbook = excelApp.ActiveWorkbook;
                var form = XmlEditForm.GetStatic();
                var controller = new XmlEditController(form);
                form.ShowForm();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
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
