using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DG2NTT.DaxDrill.DaxHelpers;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class PivotCellHelper
    {
        #region Static Members

        public static PivotCellDictionary GetPivotCellQuery(Excel.Range rngCell)
        {
            Excel.PivotTable pt = rngCell.PivotTable;
            Excel.PivotCell pc = rngCell.PivotCell; //Field values
            Excel.PivotFields pgfs = (Excel.PivotFields)(pt.PageFields);

            var pivotCellDic = new PivotCellDictionary();

            #region Filter by Single Selection

            AddSingleAxisFiltersToDic(pc, pivotCellDic);
            AddSinglePageFieldFiltersToDic(pgfs, pivotCellDic);

            #endregion

            #region Filter by Multiple Selection

            PivotTableWrapper ptw = new PivotTableWrapper(); // lazy initialization
            AddMultiplePageFieldFiltersToDic(pgfs, pivotCellDic, ptw);

            #endregion

            return pivotCellDic;
        }

        private static void AddSingleAxisFiltersToDic(Excel.PivotCell pc, PivotCellDictionary pivotCellDic)
        {
            Dictionary<string, string> singDic = pivotCellDic.SingleSelectDictionary;

            //Filter by Row and ColumnFields - note, we don't need a loop here but will use one just in case
            foreach (Excel.PivotItem pi in pc.RowItems)
            {
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                singDic.Add(pf.Name, pi.SourceName.ToString());
            }
            foreach (Excel.PivotItem pi in pc.ColumnItems)
            {
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                singDic.Add(pf.Name, pi.SourceName.ToString());
            }
        }
        
        private static Excel.PivotTable CopyAsPageInvertedPivotTable(Excel.PivotTable pt)
        {
            Excel.Application xlApp = pt.Application;
            Excel.Range rngCell = xlApp.ActiveCell;
        
            bool screenUpdating = xlApp.ScreenUpdating;
            try
            {
                xlApp.ScreenUpdating = false;
                Excel.PivotTable ptCopy = ExcelHelper.CopyAsPageInvertedPivotTable(pt);
                return ptCopy;
            }
            finally
            {
                // restore previous state
                Excel.Worksheet sheet = rngCell.Parent;
                sheet.Select();
                rngCell.Select();
                xlApp.ScreenUpdating = screenUpdating;
            }
        }

        private static void AddMultiplePageFieldFiltersToDic(Excel.PivotFields pfs, PivotCellDictionary pivotCellDic,
            PivotTableWrapper ptw)
        {
            //Filter by page field if not all items are selected
            foreach (Excel.PivotField pf in pfs)
            {
                if (ExcelHelper.IsMultiplePageItemsEnabled(pf))
                    AddMultiplePageFieldFilterToDic(pf, pivotCellDic, ptw);
            }
            
            //Clean up temp objects
            CleanUpPivotTableCopy(ptw);
        }

        private static void CleanUpPivotTableCopy(PivotTableWrapper ptw)
        {
            if (ptw == null || ptw.PivotTable == null || ptw.PivotTable == null)
                return;

            Excel.Worksheet sheet = null;
            Excel.Application xlApp = null;

            bool displayAlerts = false;

            try
            {
                sheet = ptw.PivotTable.Parent;
                xlApp = sheet.Application;
                displayAlerts = xlApp.DisplayAlerts;

                xlApp.DisplayAlerts = false;
                sheet.Delete();
            }
            finally
            {
                xlApp.DisplayAlerts = displayAlerts;
            }
        }

        private static void AddSinglePageFieldFiltersToDic(Excel.PivotFields pfs, PivotCellDictionary pivotCellDic)
        {
            //Filter by page field if not all items are selected
            foreach (Excel.PivotField pf in pfs)
            {
                if (!ExcelHelper.IsMultiplePageItemsEnabled(pf))
                    AddCurrentPageFieldFilterToDic(pf, pivotCellDic);
            }
        }

        /// <summary>
        /// This will add page filters that have multiple selection enabled to the dictionary
        /// </summary>
        /// <param name="pf">Page Field which contains multiple selections</param>
        /// <param name="pivotCellDic">Dictionary to be updated</param>
        /// <param name="ptwCopy">Object containing the copied PivotTable so that we avoid initializing it on every call</param>
        private static void AddMultiplePageFieldFilterToDic(Excel.PivotField pf, PivotCellDictionary pivotCellDic,
            PivotTableWrapper ptwCopy)
        {
            // logic
            if (ptwCopy.PivotTable == null)
            {
                Excel.PivotTable pt = pf.Parent;
                ptwCopy.PivotTable = CopyAsPageInvertedPivotTable(pt);
            }

            Excel.PivotField pfCopy = ptwCopy.PivotTable.PivotFields(pf.SourceName);
            foreach (Excel.PivotItem pi in pfCopy.VisibleItems)
            {
                pivotCellDic.AddMultiSelectItem(pfCopy.Name, pi.SourceName.ToString());
            }
        }

        private static void AddCurrentPageFieldFilterToDic(Excel.PivotField pf, PivotCellDictionary pivotCellDic)
        {
            var dicCell = pivotCellDic.SingleSelectDictionary;

            string pageName = string.Empty;

            pageName = pf.CurrentPageName; // note: throws COM exception if multiple page item selection is enabled

            bool isAllItems = true;
            isAllItems = DaxDrillParser.IsAllItems(pageName);
            if (!isAllItems)
            {
                dicCell.Add(pf.Name, pageName);
            }
        }

        public static Excel.Range CopyPivotTable(Excel.PivotTable pt)
        {
            Excel.Application XlApp = pt.Application;
            Excel.Worksheet sourceSheet = (Excel.Worksheet)pt.Parent;
            sourceSheet.Select();
            pt.PivotSelect("", Excel.XlPTSelectionMode.xlDataAndLabel, true);
            Excel.Range  sourceRange = (Excel.Range)XlApp.Selection;
            sourceRange.Copy();
            Excel.Worksheets sheets = (Excel.Worksheets)XlApp.Sheets;
            Excel.Worksheet destSheet = (Excel.Worksheet)sheets.Add();
            destSheet.Paste();
            return destSheet.Range["A1"];
        }

        private class PivotTableWrapper
        {
            private Excel.PivotTable pivotTable;
            public Excel.PivotTable PivotTable
            {
                get
                {
                    return this.pivotTable;
                }
                set
                {
                    this.pivotTable = value;
                }
            }
        }

        #endregion
    }
}
