using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DG2NTT.DaxDrill
{
    public partial class XmlEditorForm : Form
    {
        private const string AppName = "DAX Drill";

        #region Static Accessors
        public static void ShowMessage(string messageHeader, string messageDetail, string formTitle = AppName)
        {
            try
            {
                GetStatic().Text = formTitle;
                GetStatic().lblMessage.Text = messageHeader;
                GetStatic().txtStackTrace.Text = messageDetail;
                ShowForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error trying to invoke form.\n" + ex.Message + "\n" + ex.ToString(), AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Shows the static version of the form
        public static XmlEditorForm ShowForm()
        {
            try
            {
                //Show Form
                if (!GetStatic().Visible)
                {
                    GetStatic().Show();
                    GetStatic().TopMost = true;
                    //keep form on top
                }
                GetStatic().WindowState = System.Windows.Forms.FormWindowState.Normal;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error trying to show Form.\n" + ex.Message + "\n" + ex.ToString(), AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return GetStatic();
        }
        #endregion

        #region Form Initialisation via Static Accessor

        private static XmlEditorForm _form = new XmlEditorForm();

        private delegate XmlEditorForm GetFormCallBack();
        //Returns static version of the form
        public static XmlEditorForm GetStatic()
        {

            //Reinstantiate the form
            if (_form.IsDisposed)
            {
                _form = new XmlEditorForm();
            }

            if (_form.InvokeRequired)
            {
                var d = new GetFormCallBack(GetStatic);
                return (XmlEditorForm)_form.Invoke(d);
            }
            return _form;
        }

        #endregion

        private Size _originalSize;
        public XmlEditorForm()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();

            //
            // Add constructor code after the InitializeComponent() call.
            //
            _originalSize = Size;
        }

        #region Form Events

        void BtnOkClick(object sender, EventArgs e)
        {
            Hide();
        }

        private void ErrFormResize(object sender, EventArgs eventArgs)
        {
            if (WindowState == FormWindowState.Minimized) return;
            //prevent form from being resized to smaller than the minimum
            if (Size.Width < _originalSize.Width)
                Size = new Size(_originalSize.Width, Size.Height);
            if (Size.Height < _originalSize.Height)
                Size = new Size(Size.Width, _originalSize.Height);

            //resize controls
            txtStackTrace.Size = new Size(Size.Width - 41, Size.Height - 126);
            lblMessage.Size = new Size(Size.Width - 40, lblMessage.Size.Height);

            //relocate controls
            btnOk.Location = new Point(Size.Width - 104, Size.Height - 68);
        }
        #endregion
    }
}

