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
    public class XmlEditController : IDisposable
    {
        private readonly XmlEditForm xmlEditForm;
        private Excel.Workbook workbook;
        private readonly Excel.Application excelApp;
        public XmlEditController(XmlEditForm xmlEditForm)
        {
            this.xmlEditForm = xmlEditForm;

            excelApp = (Excel.Application)ExcelDnaUtil.Application;
            workbook = excelApp.ActiveWorkbook;
        }

        public void SetWorkbook(string name)
        {
            if (workbook != null) Marshal.ReleaseComObject(workbook);
            if (!string.IsNullOrEmpty(name))
                workbook = excelApp.Workbooks[name];
        }

    
        public void LoadXmlFromWorkbook()
        {
            string xmlString = ExcelHelper.ReadCustomXmlPart(workbook, xmlEditForm.NamespaceText, xmlEditForm.XpathText);
       
            //var xmls = ExcelHelper.ReadCustomXmlParts(workbook);
            //string xmlString = string.Empty;
            //foreach (string x in xmls)
            //{
            //    if (xmlString != string.Empty)
            //        xmlString += "\r\n---------\r\n";
            //    xmlString += x;
            //}
            xmlEditForm.XmlText = xmlString;
        }

        public void SaveXmlToWorkbook()
        {
            ExcelHelper.UpdateCustomXmlPart(workbook, xmlEditForm.NamespaceText, xmlEditForm.XmlText);
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        //~XmlEditController()
        //{
        //    // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //    Dispose(false);
        //}

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
