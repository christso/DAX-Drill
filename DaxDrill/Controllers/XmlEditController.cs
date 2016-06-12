using DG2NTT.DaxDrill.ExcelHelpers;
using DG2NTT.DaxDrill.UI;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DG2NTT.DaxDrill.Controllers
{
    public class XmlEditController
    {
        private readonly XmlEditForm xmlEditForm;
        public XmlEditController(XmlEditForm xmlEditForm)
        {
            this.xmlEditForm = xmlEditForm;
        }

        public void LoadXmlFromWorkbook()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelHelper.FindWorkbook(xmlEditForm.WorkbookText);
                string xmlString = ExcelHelper.ReadCustomXmlPart(workbook, xmlEditForm.NamespaceText, xmlEditForm.XpathText);
                xmlEditForm.XmlText = xmlString;
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        public void SaveXmlToWorkbook()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelHelper.FindWorkbook(xmlEditForm.WorkbookText);
                ExcelHelper.UpdateCustomXmlPart(workbook, xmlEditForm.NamespaceText, xmlEditForm.XmlText);
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }
    }
}
