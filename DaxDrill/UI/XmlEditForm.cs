using DG2NTT.DaxDrill.Controllers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DG2NTT.DaxDrill.UI
{
    public partial class XmlEditForm : Form
    {
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
        
        #endregion

        #region Static Accessors
 
        //Shows the static version of the form
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
                    xmlEditController.LoadXmlFromWorkbook();
                }

                form.WindowState = System.Windows.Forms.FormWindowState.Normal;
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

        private Size _originalSize;
        public XmlEditForm()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();

            //
            // Add constructor code after the InitializeComponent() call.
            //
            _originalSize = Size;
            xmlEditController = new XmlEditController(this);
        }

        #region Form Events

        void BtnOkClick(object sender, EventArgs e)
        {
            // save changes
            xmlEditController.SaveXmlToWorkbook();
            Hide();
        }

        private void FormResizer(object sender, EventArgs eventArgs)
        {
            if (WindowState == FormWindowState.Minimized) return;
            //prevent form from being resized to smaller than the minimum
            if (Size.Width < _originalSize.Width)
                Size = new Size(_originalSize.Width, Size.Height);
            if (Size.Height < _originalSize.Height)
                Size = new Size(Size.Width, _originalSize.Height);

            //resize controls
            txtXmlText.Size = new Size(Size.Width - 41, Size.Height - 120);

            //relocate controls
            btnOk.Location = new Point(Size.Width - 104, Size.Height - 72);
            btnCancel.Location = new Point(Size.Width - 196, Size.Height - 72);
        }
        #endregion

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}

