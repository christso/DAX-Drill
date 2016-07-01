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
        public const int PIS_LBOUND = 1;

        #region Static Members

        public static Dictionary<string, string> GetPivotCellQuery(Excel.Range rngCell)
        {
            Excel.PivotTable pt = rngCell.PivotTable;
            Excel.PivotCell pc = rngCell.PivotCell; //Field values

            Dictionary<string, string> dicCell = new Dictionary<string, string>();

            //Filter by Row and ColumnFields - note, we don't need a loop here but will use one just in case
            foreach (Excel.PivotItem pi in pc.RowItems)
            {
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                dicCell.Add(pf.Name, pi.SourceName.ToString());
            }
            foreach (Excel.PivotItem pi in pc.ColumnItems)
            {
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                dicCell.Add(pf.Name, pi.SourceName.ToString());
            }

            //Filter by page field if not all items are selected
            Excel.PivotFields pfs = (Excel.PivotFields)(pt.PageFields);

            AddOlapPageFieldFilterToDic(pfs, dicCell);

            return dicCell;
        }

        
        private static void AddOlapPageFieldFilterToDic(Excel.PivotFields pfs, Dictionary<string, string> dicCell)
        {
            //Filter by page field if not all items are selected
            foreach (Excel.PivotField pf in pfs)
            {
                if (ExcelHelper.IsMultiplePageItemsEnabled(pf))
                    continue;

                string pageName = string.Empty;

                pageName = pf.CurrentPageName; // note: throws exception if multiple page item selection is enabled

                bool isAllItems = true;
                isAllItems = DaxDrillParser.IsAllItems(pageName);
                if (!isAllItems)
                {
                    dicCell.Add(pf.Name, pageName);
                }
            }
        }

        public static Excel.Range CopyPivotTable(Excel.PivotTable pt)
        {
            Excel.Application XlApp = null;
            Excel.Worksheet sourceSheet = null;
            Excel.Range sourceRange = null;
            Excel.Worksheet destSheet = null;
            Excel.Worksheets sheets = null;

            XlApp = pt.Application;
            sourceSheet = (Excel.Worksheet)pt.Parent;
            sourceSheet.Select();
            pt.PivotSelect("", Excel.XlPTSelectionMode.xlDataAndLabel, true);
            sourceRange = (Excel.Range)XlApp.Selection;
            sourceRange.Copy();
            sheets = (Excel.Worksheets)XlApp.Sheets;
            destSheet = (Excel.Worksheet)sheets.Add();
            destSheet.Paste();
            return destSheet.Range["A1"];
        }

        #endregion
    }
}
