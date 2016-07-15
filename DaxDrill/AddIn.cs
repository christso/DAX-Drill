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
using DG2NTT.DaxDrill.Helpers;
using System.Collections;
using Office = Microsoft.Office.Core;

namespace DG2NTT.DaxDrill
{
    public class AddIn : IExcelAddIn
    {
        private static Excel.Application xlApp
        {
            get
            {
                return (Excel.Application)ExcelDnaUtil.Application;
            }
        }

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

        private void SheetBeforeDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            try
            {
                Cancel = false;
                Excel.Range rngCell = xlApp.ActiveCell;
                if (!ExcelHelper.IsPivotDataCell(rngCell)) return;
                if (!QueryClient.IsDrillThroughEnabled(rngCell)) return;
                DrillThrough();
                Cancel = true;
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        private void XlApp_WorkbookDeactivate(Excel.Workbook Wb)
        {
            if (Wb.Application.Workbooks.Count == 1)
            {
                //uncomment below if you want to clean up using GC
                CleanUp();

                //uncomment below if you need to kill the Excel process
                //Process.GetCurrentProcess().Kill();
            }
        }

        private static void CleanUp()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            Excel.Range rngCell = xlApp.ActiveCell;
            if (!ExcelHelper.IsPivotDataCell(rngCell)) return;

            // set up connection
            var queryClient = new QueryClient(rngCell);
            var connString = ExcelHelper.GetConnectionString(rngCell);
            var commandText = queryClient.GetDAXQuery(connString);
            var daxClient = new DaxClient();
            var cnnStringBuilder = new TabularConnectionStringBuilder(connString);
            var cnn = new ADOMD.AdomdConnection(cnnStringBuilder.StrippedConnectionString);

            // create sheet
            Excel.Sheets sheets = xlApp.Sheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Add();

            // show message to user we are retrieving records
            Excel.Range rngHead = sheet.Range["A1"];
            int maxDrillThroughRecords = ExcelHelper.GetMaxDrillthroughRecords(rngCell);
            rngHead.Value2 = string.Format("Retrieving TOP {0} records", 
                maxDrillThroughRecords);
                
            // retrieve result
            var dtResult = daxClient.ExecuteTable(commandText, cnn);

            // output result to sheet
            Excel.Range rngOut = sheet.Range["A3"];
            ExcelHelper.FillRange(dtResult, rngOut);
            ExcelHelper.FormatRange(dtResult, rngOut);
            rngHead.Value2 = string.Format("Retrieved TOP {0} records", maxDrillThroughRecords);
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "DAX Query")]
        public static void DrillThroughQuery()
        {
            try
            {
                Excel.Range rngCell = xlApp.ActiveCell;
                if (!ExcelHelper.IsPivotDataCell(rngCell)) return;

                // generate command
                var queryClient = new QueryClient(rngCell);
                var commandText = queryClient.GetDAXQuery();
                MsgForm.ShowMessage("DAX Query", commandText);
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "XML Metadata")]
        public static void ShowMetadataEditor()
        {
            try
            {
                var form = XmlEditForm.GetStatic();
                var controller = new XmlEditController(form);
                form.ShowForm();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        [ExcelCommand(MenuName = "&DAX Drill", MenuText = "About")]
        public static void About()
        {
            try
            {
                AboutBox.ShowForm();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }
    }
}
