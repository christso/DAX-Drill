using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DG2NTT.DaxDrill.ExcelHelpers
{
    public class PivotTableWrapper
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
}
