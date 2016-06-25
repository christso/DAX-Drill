using DG2NTT.DaxDrill.Logic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using DG2NTT.DaxDrill.ExcelHelpers;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace DG2NTT.DaxDrill.UI
{
    public partial class XmlEditForm : Form
    {
        public XmlEditForm()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();

            //
            // Add constructor code after the InitializeComponent() call.
            //
            xmlEditController = new XmlEditController(this);
        }

        private const string AppName = Constants.AppName;
        private readonly XmlEditController xmlEditController;

        #region Public Members

        public string XmlText
        {
            get { return txtXmlText.Text; }
            set { txtXmlText.Text = value;  }
        }

        public string FormTitle
        {
            get { return this.Text; }
            set { this.Text = value;  }
        }

        public string NamespaceText
        {
            get { return cbNamespace.Text; }
            set { cbNamespace.Text = value; }
        }

        public string XpathText
        {
            get { return txtXpath.Text;  }
            set { txtXpath.Text = value; }
        }

        public string WorkbookText
        {
            get { return cbWorkbooks.Text; }
            set { cbWorkbooks.Text = value; }
        }

        public void ShowForm()
        {
            try
            {
                var form = this;

                //Show Form
                if (!form.Visible)
                {
                    form.Show();
                    form.TopMost = true; //keep form on top
                }

                form.WindowState = System.Windows.Forms.FormWindowState.Normal;

                RefreshWorkbooksControl();
                SelectActiveWorkbook();
                RefreshNamespaceControl();
                SelectDefaultNamespace();
                txtXpath.Text = "x:*";
                xmlEditController.LoadXmlFromWorkbook();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        public void RefreshNamespaceControl()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = ExcelHelper.FindWorkbook(WorkbookText);
                RefreshNamespaceControl(workbook);
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        public void RefreshNamespaceControl(Excel.Workbook workbook)
        {
            var nsList = ExcelHelper.ListXmlNamespaces(workbook);
            cbNamespace.Items.Clear();
            cbNamespace.Items.AddRange(nsList.ToArray());
            
        }

        public void SelectDefaultNamespace()
        {
            NamespaceText = Constants.DaxDrillXmlSchemaSpace;
        }

        public void SelectActiveWorkbook()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                workbook = excelApp.ActiveWorkbook;
                cbWorkbooks.Text = workbook.Name;
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }

        public void RefreshWorkbooksControl()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = (Excel.Application)ExcelDnaUtil.Application;
                
                workbook = excelApp.ActiveWorkbook;
                var wbList = ExcelHelper.ListWorkbooks(excelApp);
                cbWorkbooks.Items.Clear();
                cbWorkbooks.Items.AddRange(wbList.ToArray());
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
            }
        }


        #endregion

        #region Form Events

        void BtnSaveClick(object sender, EventArgs e)
        {
            try
            {
                // save changes
                xmlEditController.SaveXmlToWorkbook();
                MessageBox.Show(string.Format("XML saved to namespace '{0}'", this.NamespaceText));
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            try
            {
                xmlEditController.LoadXmlFromWorkbook();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        #endregion

        #region Form Initialisation via Static Accessor

        private static XmlEditForm _form = new XmlEditForm();

        private delegate XmlEditForm GetFormCallBack();
        //Returns static version of the form
        public static XmlEditForm GetStatic()
        {

            //Reinstantiate the form
            if (_form.IsDisposed)
            {
                _form = new XmlEditForm();
            }

            if (_form.InvokeRequired)
            {
                var d = new GetFormCallBack(GetStatic);
                return (XmlEditForm)_form.Invoke(d);
            }
            return _form;
        }

        #endregion

        private void cbWorkbooks_SelectedValueChanged(object sender, EventArgs e)
        {
            RefreshNamespaceControl();
        }
    }
}

