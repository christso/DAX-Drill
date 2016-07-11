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

            AddSingleAxisFiltersToDic(pc, pivotCellDic);
            AddSinglePageFieldFiltersToDic(pgfs, pivotCellDic);
            AddMultiplePageFieldFilterToDic(pt, pivotCellDic);

            return pivotCellDic;
        }

        private static void AddMultiplePageFieldFilterToDic(Excel.PivotTable pt, PivotCellDictionary pivotCellDic)
        {
            var mdxString = pt.MDX;
            var daxFilterList = DaxDrillParser.ConvertExcelMdxToDaxFilter(mdxString);
            DaxDrillParser.ConvertDaxFilterListToDictionary(daxFilterList, pivotCellDic.MultiSelectDictionary);
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
        
        private static void AddSinglePageFieldFiltersToDic(Excel.PivotFields pfs, PivotCellDictionary pivotCellDic)
        {
            //Filter by page field if not all items are selected
            foreach (Excel.PivotField pf in pfs)
            {
                var dicCell = pivotCellDic.SingleSelectDictionary;

                string pageName = pf.DataRange.Value2;
                if (pageName != "All" && pageName != "(Multiple Items)")
                {
                    dicCell.Add(pf.Name, pf.CubeField.Name + ".&[" + pageName + "]");
                }
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

        #endregion
    }
}
