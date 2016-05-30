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
    public class ExcelHelper
    {
        private Excel.Application excelApp;
        public ExcelHelper(Excel.Application excelApp)
        {
            this.excelApp = excelApp;
        }

        public void GetDAXQuery(Excel.Range rngCell)
        {
            var queryDic = PivotTableHelper.GetPivotCellQuery(rngCell);

        }

        public void FillRange(System.Data.DataTable dataTable, Excel.Range rngOutput)
        {
            Excel.Worksheet sheet = excelApp.ActiveSheet;
            Excel.Range rng = null;
            const int boundToSizeFactor = 1;
            const int rowBoundIndex = 0;
            const int columnBoundIndex = 1;

            try
            {

                object[,] arr = Utils.CreateArray(dataTable);
                rng = rngOutput.Resize[arr.GetUpperBound(rowBoundIndex) + boundToSizeFactor,
                    arr.GetUpperBound(columnBoundIndex) + boundToSizeFactor];
                rng.Value2 = arr;

            }
            finally
            {
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (rng != null) Marshal.ReleaseComObject(rng);
            }
        }

    }
}
