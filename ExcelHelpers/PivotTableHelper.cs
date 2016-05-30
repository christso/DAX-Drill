using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class PivotTableHelper
    {
        #region Static Members
        public static Dictionary<string, string> GetPivotCellQuery(Excel.Range rngCell)
        {
            Excel.Application XlApp = null;
            //Field values
            Excel.PivotCell pc = null;
            //Aggregate value the user is drilling through
            Excel.PivotTable pt = null;

            try
            {
                XlApp = rngCell.Application;
                pt = XlApp.ActiveCell.PivotTable;
                pc = XlApp.ActiveCell.PivotCell;

                Dictionary<string, string> dicCell = new Dictionary<string, string>();

                //Filter by Row and ColumnFields - note, we don't need a loop here but will use one just in case
                for (int i = 0; i < pc.RowItems.Count; i++)
                {
                    Excel.PivotItem pi = pc.RowItems[i];
                    Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                    dicCell.Add(pf.Name, pi.SourceName.ToString());

                    if (pi != null) Marshal.ReleaseComObject(pi);
                    if (pf != null) Marshal.ReleaseComObject(pf);
                }
                for (int i = 0; i < pc.ColumnItems.Count; i++)
                {
                    Excel.PivotItem pi = pc.ColumnItems[i];
                    Excel.PivotField pf = (Excel.PivotField)pi.Parent;
                    dicCell.Add(pf.Name, pi.SourceName.ToString());

                    if (pi != null) Marshal.ReleaseComObject(pi);
                    if (pf != null) Marshal.ReleaseComObject(pf);
                }

                //Filter by page field if not all items are selected
                foreach (Excel.PivotField pf in (Excel.PivotFields)(pt.PageFields))
                {
                    var currentPage = (Excel.PivotItem)pf.CurrentPage;
                    if (currentPage.Name != "(All)")
                    {
                        Excel.PivotItem pi = (Excel.PivotItem)pf.CurrentPage;
                        dicCell.Add(pf.Name, pi.SourceName.ToString());
                    }
                    if (currentPage != null) Marshal.ReleaseComObject(currentPage);
                }

                return dicCell;
            }
            finally
            {
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
                if (pt != null) Marshal.ReleaseComObject(pt);
                if (pc != null) Marshal.ReleaseComObject(pc);
            }
        }


        public static Dictionary<string, List<string>> GetPivotCellHiddenQuery(Excel.Range rngCell)
        {
            Excel.Application XlApp = null;
            Excel.PivotTable pt = null;
            Excel.PivotFields pfs = null;
            Excel.PivotFields rfs = null;

            try
            {
                XlApp = rngCell.Application;
                pfs = (Excel.PivotFields)pt.PageFields;

                //Field names
                pt = XlApp.ActiveCell.PivotTable;

                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();

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

                //Row fields

                rfs = (Excel.PivotFields)pt.RowFields;

                for (int i = 0; i < rfs.Count; i++)
                {
                    var pf = (Excel.PivotField)rfs.Item(i);

                    //Get hidden items for page fields where not all items are visible
                    List<string> hiddenItems = HiddenPivotFieldItems(pf);

                    if (hiddenItems.Count > 0)
                    {
                        //Add list to dictionary
                        dic.Add(pf.SourceName, hiddenItems);
                    }
                    if (pf != null) Marshal.ReleaseComObject(pf);
                }

                //Column Fields

                foreach (Excel.PivotField pf in (Excel.PivotFields)pt.ColumnFields)
                {

                    //Get hidden items for page fields where not all items are visible

                    List<string> sList = HiddenPivotFieldItems(pf);

                    if (sList.Count > 0)
                    {
                        //Add list to dictionary
                        dic.Add(pf.SourceName, sList);
                    }

                }

                return dic;
            }
            finally
            {

            }
        }

        public static Excel.Range CopyPivotTable(Excel.PivotTable pt)
        {
            Excel.Application XlApp = null;
            Excel.Worksheet sourceSheet = null;
            Excel.Range sourceRange = null;
            Excel.Worksheet destSheet = null;

            try
            {
                XlApp = pt.Application;
                sourceSheet = (Excel.Worksheet)pt.Parent;

                sourceSheet.Select();
                pt.PivotSelect("", Excel.XlPTSelectionMode.xlDataAndLabel, true);
                sourceRange = (Excel.Range)XlApp.Selection;
                sourceRange.Copy();
                destSheet = (Excel.Worksheet)XlApp.Sheets.Add();
                destSheet.Paste();
                return destSheet.Range["A1"];
            }
            finally
            {
                if (XlApp != null) Marshal.ReleaseComObject(XlApp);
                if (sourceSheet != null) Marshal.ReleaseComObject(sourceSheet);
                if (sourceRange != null) Marshal.ReleaseComObject(sourceRange);
                if (destSheet != null) Marshal.ReleaseComObject(destSheet);
            }
        }

        //Returns list of hidden items of a pivot field (including pagefield)
        private static List<string> HiddenPivotFieldItems(Excel.PivotField pf)
        {

            Excel.PivotTable ptCopy;
            Excel.PivotField pfCopy;
            Excel.PivotTable pt;
            Excel.Worksheet wsCopy = null;

            //Change page field to rowfield

            if (pf.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
            {
                pt = (Excel.PivotTable)pf.Parent;
                Excel.Range rngCopy = CopyPivotTable(pt);
                ptCopy = rngCopy.PivotTable;

                Excel.PivotField _pf = (Excel.PivotField)ptCopy.PivotFields(pf.Name);
                try
                {
                    _pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }
                catch
                {
                    pt.RefreshTable();
                }
                _pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                _pf.Position = 1;

                pfCopy = (Excel.PivotField)ptCopy.PivotFields(_pf.Name);
                wsCopy = (Excel.Worksheet)rngCopy.Worksheet;
            }
            else
            {
                pfCopy = pf;

            }

            List<string> result = HiddenNonPageFieldItems(pfCopy);

            var XlApp = (Excel.Application)pf.Application;
            //Delete worksheet where the pivot table copy resides
            bool DisplayAlerts = XlApp.DisplayAlerts;
            try
            {
                XlApp.DisplayAlerts = false;
                if ((wsCopy != null))
                {
                    wsCopy.Delete();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                XlApp.DisplayAlerts = DisplayAlerts;
            }

            return result;

        }

        //Returns list of hidden items of a pivot field (except for pagefield due to Excel bug)
        private static List<string> HiddenNonPageFieldItems(Excel.PivotField pf)
        {

            List<string> sList = new List<string>();

            //Add hidden items to list
            Excel.PivotItems pis = (Excel.PivotItems)pf.HiddenItems;

            foreach (Excel.PivotItem pi in pis)
            {
                sList.Add((string)pi.SourceName);
            }

            return sList;
        }

        #endregion
    }
}
