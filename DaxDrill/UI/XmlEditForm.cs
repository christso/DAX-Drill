﻿using DG2NTT.DaxDrill.Controllers;
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

                RefreshControls();
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        public void RefreshControls()
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
                cbWorkbooks.Text = workbook.Name;

                var nsList = ExcelHelper.ListXmlNamespaces(workbook);
                cbNamespace.Items.Clear();
                cbNamespace.Items.AddRange(nsList.ToArray());
                cbNamespace.Text = Constants.DaxDrillXmlSchemaSpace;

                txtXpath.Text = "x:*";

                xmlEditController.LoadXmlFromWorkbook();
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
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

        private void btnRefresh_Click(object sender, EventArgs e)
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

    }
}
