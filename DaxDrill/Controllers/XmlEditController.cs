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
        private readonly Excel.Workbook workbook;
        private readonly Excel.Application excelApp;
        public XmlEditController(XmlEditForm xmlEditForm)
        {
            this.xmlEditForm = xmlEditForm;

            excelApp = (Excel.Application)ExcelDnaUtil.Application;
            workbook = excelApp.ActiveWorkbook;
        }

        public void LoadXmlFromWorkbook()
        {
            string xml = ExcelHelper.ReadCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, "/x:table");
            if (string.IsNullOrEmpty(xml)) ExcelHelper.ReadCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, "/x:columns");
            xmlEditForm.XmlText = xml;
        }

        public void SaveXmlToWorkbook()
        {
            //ExcelHelper.UpdateCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, xmlEditForm.XmlText);

            var xmlString =
@"<?xml version=""1.0"" encoding=""utf-8\"" ?>
<table id=""Usage"" connection_id=""localhost Roaming Model"" xmlns=""{0}"">
	<columns>
	   <column>
		  <name>Call Type</name>
		  <expression>Usage[Call Type]</expression>
	   </column>
	   <column>
		  <name>Call Type Description</name>
		  <expression>Usage[Call Type Description]</expression>
	   </column>
	   <column>
		  <name>Gross Billed</name>
		  <expression>Usage[Gross Billed]</expression>
	   </column>
	</columns>
</table>".Replace("{0}", Constants.DaxDrillXmlSchemaSpace);

            ExcelHelper.UpdateCustomXmlPart(workbook, Constants.DaxDrillXmlSchemaSpace, xmlString);
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
