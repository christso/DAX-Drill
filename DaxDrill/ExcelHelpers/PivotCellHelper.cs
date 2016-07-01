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
            Excel.Application XlApp = null;
            //Field values
            Excel.PivotCell pc = null;
            //Aggregate value the user is drilling through
            Excel.PivotTable pt = null;
            Excel.PivotFields pfs = null;
            Excel.PivotCache cache = null;

            XlApp = rngCell.Application;
            pt = rngCell.PivotTable;
            pc = rngCell.PivotCell;

            Dictionary<string, string> dicCell = new Dictionary<string, string>();

            //Filter by Row and ColumnFields - note, we don't need a loop here but will use one just in case
            for (int i = PIS_LBOUND; i < pc.RowItems.Count + PIS_LBOUND; i++)
            {
                Excel.PivotItem pi = pc.RowItems[i];
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                dicCell.Add(pf.Name, pi.SourceName.ToString());
            }
            for (int i = PIS_LBOUND; i < pc.ColumnItems.Count + PIS_LBOUND; i++)
            {
                Excel.PivotItem pi = pc.ColumnItems[i];
                Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                dicCell.Add(pf.Name, pi.SourceName.ToString());
            }

            //Filter by page field if not all items are selected
            pfs = (Excel.PivotFields)(pt.PageFields);
            cache = pt.PivotCache();
            if (cache.OLAP)
                AddOlapPageFieldFilterToDic(pfs, dicCell);
            else
                AddPageFieldFilterToDic(pfs, dicCell);
                
            return dicCell;
        }

        private static void AddOlapPageFieldFilterToDic(Excel.PivotFields pfs, Dictionary<string, string> dicCell)
        {
            //Filter by page field if not all items are selected
            for (int i = PIS_LBOUND; i < pfs.Count + PIS_LBOUND; i++)
            {
                Excel.PivotField pf = pfs.Item(i);
                bool isAllItems = true;
                string pageName = string.Empty;

                try
                {
                    pageName = pf.CurrentPageName; // error maybe thrown
                    isAllItems = DaxDrillParser.IsAllItems(pageName);

                    if (!isAllItems)
                    {
                        dicCell.Add(pf.Name, pageName);
                    }
                }
                catch (COMException ex)
                {
                    // exception is thrown by Excel if multiple item selection is enabled
                    // TODO: create filters for multiple item selection
                }
            }
        }

        private static void AddPageFieldFilterToDic(Excel.PivotFields pfs, Dictionary<string, string> dicCell)
        {
            //Filter by page field if not all items are selected
            for (int i = PIS_LBOUND; i < pfs.Count + PIS_LBOUND; i++)
            {
                Excel.PivotField pf = pfs.Item(i);
                var currentPage = (Excel.PivotItem)pf.CurrentPage;
                if (currentPage.Name != "(All)")
                {
                    Excel.PivotItem pi = (Excel.PivotItem)pf.CurrentPage;
                    dicCell.Add(pf.Name, pi.SourceName.ToString());
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
