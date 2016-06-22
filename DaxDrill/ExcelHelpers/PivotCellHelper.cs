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

            try
            {
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

                    if (pi != null) Marshal.ReleaseComObject(pi);
                    if (pf != null) Marshal.ReleaseComObject(pf);
                }
                for (int i = PIS_LBOUND; i < pc.ColumnItems.Count + PIS_LBOUND; i++)
                {
                    Excel.PivotItem pi = pc.ColumnItems[i];
                    Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                    dicCell.Add(pf.Name, pi.SourceName.ToString());

                    if (pi != null) Marshal.ReleaseComObject(pi);
                    if (pf != null) Marshal.ReleaseComObject(pf);
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
            finally
            {
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (pc != null) Marshal.ReleaseComObject(pc);
                if (pfs != null) Marshal.ReleaseComObject(pfs);
                if (cache != null) Marshal.ReleaseComObject(cache);
            }
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
                }
                finally
                {
                    if (pf != null) Marshal.ReleaseComObject(pf);
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
                if (currentPage != null) Marshal.ReleaseComObject(currentPage);
                if (pf != null) Marshal.ReleaseComObject(pf);
            }
        }

        public static Dictionary<string, List<string>> GetPivotCellHiddenQuery(Excel.Range rngCell)
        {
            Excel.Application XlApp = null;
            Excel.PivotTable pt = null;
            Excel.PivotFields pfs = null;
            Excel.PivotFields rfs = null;
            Excel.PivotFields cfs = null;

            try
            {
                XlApp = rngCell.Application;
                pt = XlApp.ActiveCell.PivotTable;
                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();

                // page fields
                pfs = (Excel.PivotFields)pt.PageFields;
                AddHiddenItemsToDictionary(pfs, dic);

                // row fields
                rfs = (Excel.PivotFields)pt.RowFields;
                AddHiddenItemsToDictionary(rfs, dic);

                //Column Fields
                cfs = (Excel.PivotFields)pt.ColumnFields;
                AddHiddenItemsToDictionary(cfs, dic);

                return dic;
            }
            finally
            {
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (pfs != null) Marshal.ReleaseComObject(pfs);
                if (rfs != null) Marshal.ReleaseComObject(rfs);
                if (cfs != null) Marshal.ReleaseComObject(cfs);
            }
        }

        private static void AddHiddenItemsToDictionary(Excel.PivotFields pfs, Dictionary<string, List<string>> dic)
        {
            for (int i = 0; i < pfs.Count; i++)
            {
                var pf = (Excel.PivotField)pfs.Item(i);
                //Get hidden items for page fields where not all items are visible
                List<string> hiddenItems = HiddenPivotFieldItems(pf);

                if (hiddenItems.Count > 0)
                {
                    //Add list to dictionary
                    dic.Add(pf.SourceName, hiddenItems);
                }
                if (pf != null) Marshal.ReleaseComObject(pf);
            }
        }

        public static Excel.Range CopyPivotTable(Excel.PivotTable pt)
        {
            Excel.Application XlApp = null;
            Excel.Worksheet sourceSheet = null;
            Excel.Range sourceRange = null;
            Excel.Worksheet destSheet = null;
            Excel.Worksheets sheets = null;

            try
            {
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
            finally
            {
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
                if (sourceSheet != null) Marshal.ReleaseComObject(sourceSheet);
                if (sourceRange != null) Marshal.ReleaseComObject(sourceRange);
                if (destSheet != null) Marshal.ReleaseComObject(destSheet);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }
        }

        //Returns list of hidden items of a pivot field (including pagefield)
        private static List<string> HiddenPivotFieldItems(Excel.PivotField pf)
        {
            Excel.PivotTable ptCopy = null;
            Excel.PivotField pfCopy = null;
            Excel.PivotTable pt = null;
            Excel.Worksheet wsCopy = null;
            Excel.Application XlApp = null;

            try
            {
                XlApp = (Excel.Application)pf.Application;

                #region Change page field to rowfield
                // This will copy the pivot table into a new temporary worksheet
                if (pf.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
                {
                    pt = (Excel.PivotTable)pf.Parent;
                    Excel.Range rngCopy = CopyPivotTable(pt);
                    ptCopy = rngCopy.PivotTable;

                    Excel.PivotField rf = (Excel.PivotField)ptCopy.PivotFields(pf.Name);
                    try
                    {
                        rf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    }
                    catch
                    {
                        pt.RefreshTable();
                    }
                    rf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    rf.Position = 1;

                    pfCopy = (Excel.PivotField)ptCopy.PivotFields(rf.Name);
                    wsCopy = (Excel.Worksheet)rngCopy.Worksheet;
                }
                else
                {
                    pfCopy = pf;

                }
                #endregion

                // get hidden items for non-page fields
                List<string> result = HiddenNonPageFieldItems(pfCopy);

                #region Delete temporary worksheet
                bool DisplayAlerts = XlApp.DisplayAlerts;
                try
                {
                    XlApp.DisplayAlerts = false;
                    if ((wsCopy != null))
                    {
                        wsCopy.Delete();
                    }
                }
                finally
                {
                    XlApp.DisplayAlerts = DisplayAlerts;
                }
                #endregion

                return result;
            }
            finally
            {
                if (ptCopy != null) Marshal.ReleaseComObject(ptCopy);
                if (pfCopy != null) Marshal.ReleaseComObject(pfCopy);
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (wsCopy != null) Marshal.ReleaseComObject(wsCopy);
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
            }
        }

        //Returns list of hidden items of a pivot field (except for pagefield due to Excel bug)
        private static List<string> HiddenNonPageFieldItems(Excel.PivotField pf)
        {
            Excel.PivotItems pis = null;

            try
            {
                List<string> result = new List<string>();

                //Add hidden items to list
                pis = (Excel.PivotItems)pf.HiddenItems;

                for (int i = 0; i < pis.Count; i++)
                {
                    Excel.PivotItem pi = pis.Item(i);
                    result.Add((string)pi.SourceName);
                    if (pi != null) Marshal.ReleaseComObject(pi);
                }

                return result;
            }
            finally
            {
                if (pis != null) Marshal.ReleaseComObject(pis);
            }
        }

        #endregion
    }
}
