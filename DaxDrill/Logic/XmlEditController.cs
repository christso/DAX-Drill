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

namespace DG2NTT.DaxDrill.Logic
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
            workbook = ExcelHelper.FindWorkbook(xmlEditForm.WorkbookText);
            string xmlString = ExcelHelper.ReadCustomXmlNode(workbook, xmlEditForm.NamespaceText, xmlEditForm.XpathText);
            xmlEditForm.XmlText = xmlString;
        }

        public void SaveXmlToWorkbook()
        {
            Excel.Workbook workbook = ExcelHelper.FindWorkbook(xmlEditForm.WorkbookText);

            // save XML
            if (xmlEditForm.XpathText == "x:*" || string.IsNullOrWhiteSpace(xmlEditForm.XpathText))
                ExcelHelper.UpdateCustomXmlPart(workbook, xmlEditForm.NamespaceText, xmlEditForm.XmlText);
            else
                ExcelHelper.UpdateCustomXmlNode(workbook, xmlEditForm.NamespaceText, xmlEditForm.XmlText, xmlEditForm.XpathText);
        }

        public void InitializeXml()
        {
            Excel.Workbook workbook = ExcelHelper.FindWorkbook(xmlEditForm.WorkbookText);

            // create daxdrill node template
            ExcelHelper.AddCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace,
                string.Format("<daxdrill xmlns=\"{0}\"></daxdrill>", Constants.DaxDrillXmlSchemaSpace));
        }
    }
}
