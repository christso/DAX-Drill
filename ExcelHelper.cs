using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace DG2NTT.DaxDrill
{
    public class ExcelHelper
    {
        private Excel.Application excelApp;
        public ExcelHelper(Excel.Application excelApp)
        {
            this.excelApp = excelApp;
        }

        public void FillRange(System.Data.DataTable dataTable, Excel.Range rngOutput)
        {
            Excel.Worksheet sheet = excelApp.ActiveSheet;

            object[,] arr = CreateArray(dataTable);
            Excel.Range rng = rngOutput.Resize[arr.GetUpperBound(0) + 1, arr.GetUpperBound(1) + 1];
            rng.Value2 = arr;

            if (sheet != null) Marshal.ReleaseComObject(sheet);
            if (rng != null) Marshal.ReleaseComObject(rng);
        }

        public object[,] CreateArray(System.Data.DataTable dataTable)
        {
            var rowCount = dataTable.Rows.Count;
            var columnCount = dataTable.Columns.Count;
            object[,] result = new object[rowCount, columnCount];

            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    result[r, c] = dataTable.Rows[r][c];
                }
            }

            return result;
        }

        public static Dictionary<string, string> GetPivotCellQuery(Excel.Range rngCell)
        {
            Excel.Application XlApp = rngCell.Application;

            //Field values
            Excel.PivotCell pc = null;
            //Aggregate value the user is drilling through
            Excel.PivotTable pt = null;

            pt = XlApp.ActiveCell.PivotTable;
            pc = XlApp.ActiveCell.PivotCell;

            Dictionary<string, string> dicCell = new Dictionary<string, string>();

            //Filter by Row and ColumnFields - not, we don't need a loop here but will use one just in case
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
            foreach (Excel.PivotField pf in (Excel.PivotFields)(pt.PageFields))
            {
                if (((Excel.PivotItem)pf.CurrentPage).Name != "(All)")
                {
                    Excel.PivotItem pi = (Excel.PivotItem)pf.CurrentPage;
                    dicCell.Add(pf.Name, pi.SourceName.ToString());
                }
            }

            return dicCell;
        }
    }
}
