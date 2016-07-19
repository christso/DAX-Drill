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
            try
            {
                var xlApp = (Excel.Application)ExcelDnaUtil.Application;
                Excel.PivotField pf = xlApp.ActiveCell.PivotField;
                ExcelHelper.SetPivotFieldPage(pf, pivotFieldPageEditForm.PageItemValue);
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        public void GetPivotFieldPage()
        {
            try
            {
                var xlApp = (Excel.Application)ExcelDnaUtil.Application;
                var rngCell = (Excel.Range)xlApp.ActiveCell;
                var pf = (Excel.PivotField)rngCell.PivotField;
                if (ExcelHelper.IsPivotPageField(rngCell))
                    pivotFieldPageEditForm.PageItemValue = DaxHelpers.DaxDrillParser.GetValueFromPivotItem(pf.CurrentPageName);
                else
                    pivotFieldPageEditForm.PageItemValue = string.Empty;
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }
    }
}
