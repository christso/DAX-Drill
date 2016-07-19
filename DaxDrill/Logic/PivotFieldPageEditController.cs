using DG2NTT.DaxDrill.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DG2NTT.DaxDrill.ExcelHelpers;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace DG2NTT.DaxDrill.Logic
{
    public class PivotFieldPageEditController
    {
        private readonly PivotFieldPageEditForm pivotFieldPageEditForm;

        public PivotFieldPageEditController(PivotFieldPageEditForm pivotFieldPageEditForm)
        {
            this.pivotFieldPageEditForm = pivotFieldPageEditForm;
        }

        public void SetPivotFieldPage()
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            Excel.PivotField pf = xlApp.ActiveCell.PivotField;
            ExcelHelper.SetPivotFieldPage(pf, pivotFieldPageEditForm.PageItemValue);
        }
    }
}
