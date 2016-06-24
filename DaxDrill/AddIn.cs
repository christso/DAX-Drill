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
        private static Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;

        public void AutoClose()
        {
            xlApp.WorkbookDeactivate -= XlApp_WorkbookDeactivate;
            xlApp.SheetBeforeDoubleClick -= SheetBeforeDoubleClick;
        }

        public void AutoOpen()
        {
            xlApp.WorkbookDeactivate += XlApp_WorkbookDeactivate;
            xlApp.SheetBeforeDoubleClick += SheetBeforeDoubleClick;
        }

        // double click stops working after 15 times
        private void SheetBeforeDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            try
            {
                DrillThrough();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
            Cancel = true;
        }

        // kill Excel process in case objects are not properly released
        private void XlApp_WorkbookDeactivate(Excel.Workbook Wb)
        {
            if (Wb.Application.Workbooks.Count == 1)
            {
                if (xlApp != null) Marshal.ReleaseComObject(xlApp);
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
            Excel.Range rngHead = null;
            Excel.Range rngOut = null;
            Excel.Range rngCell = null;

            try
            {
                rngCell = xlApp.ActiveCell;

                // create sheet
                sheets = xlApp.Sheets;
                sheet = (Excel.Worksheet)sheets.Add();

                rngHead = sheet.Range["A1"];
                int maxDrillThroughRecords = ExcelHelper.GetMaxDrillthroughRecords(rngCell);
                rngHead.Value2 = string.Format("Retrieving TOP {0} records", 
                    maxDrillThroughRecords);
                
                // set up connection
                var queryClient = new QueryClient(rngCell);
                var connString = ExcelHelper.GetConnectionString(rngCell);
                var commandText = queryClient.GetDAXQuery(connString, rngCell);
                var daxClient = new DaxClient();
                var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
                var cnn = new ADOMD.AdomdConnection(cnnStringBuilder.StrippedConnectionString);
                var dtResult = daxClient.ExecuteTable(commandText, cnn);

                // output result to sheet
                rngOut = sheet.Range["A3"];
                ExcelHelper.FillRange(dtResult, rngOut);
                rngHead.Value2 = string.Format("Retrieved TOP {0} records", maxDrillThroughRecords);
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rngOut != null) Marshal.ReleaseComObject(rngOut);
                if (rngHead != null) Marshal.ReleaseComObject(rngHead);
                if (rngCell != null) Marshal.ReleaseComObject(rngCell);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DAX Query")]
        public static void DrillThroughQuery()
        {
            Excel.Range rngCell = null;
            Excel.Workbook workbook = null;

            try
            {
                // XML configuration
                workbook = xlApp.ActiveWorkbook;
                string xml = ExcelHelper.ReadCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, "/x:columns");
                
                // generate command
                rngCell = xlApp.ActiveCell;
                var queryClient = new QueryClient(rngCell);
                var commandText = queryClient.GetDAXQuery(rngCell);
                MsgForm.ShowMessage("DAX Query", commandText);
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
            finally
            {
                if (rngCell != null) Marshal.ReleaseComObject(rngCell);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "XML Metadata")]
        public static void ShowMetadataEditor()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = xlApp.ActiveWorkbook;
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
