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

namespace DG2NTT.DaxDrill.UI
{
    public partial class PivotFieldPageEditForm : Form
    {
        public PivotFieldPageEditForm()
        {
            InitializeComponent();
            this.pivotFieldPageEditController = new PivotFieldPageEditController(this);
        }

        private readonly PivotFieldPageEditController pivotFieldPageEditController;

        private void btnOk_Click(object sender, EventArgs e)
        {
            pivotFieldPageEditController.SetPivotFieldPage();
        }
        
        public string PageItemValue
        {
            get
            {
                return this.txtPageItemValue.Text;
            }
            set
            {
                this.txtPageItemValue.Text = value;
            }
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
            }
            catch (Exception ex)
            {
                MsgForm.ShowMessage(ex);
            }
        }

        #region Form Initialisation via Static Accessor

        private static PivotFieldPageEditForm _form = new PivotFieldPageEditForm();

        private delegate PivotFieldPageEditForm GetFormCallBack();
        //Returns static version of the form
        public static PivotFieldPageEditForm GetStatic()
        {

            //Reinstantiate the form
            if (_form.IsDisposed)
            {
                _form = new PivotFieldPageEditForm();
            }

            if (_form.InvokeRequired)
            {
                var d = new GetFormCallBack(GetStatic);
                return (PivotFieldPageEditForm)_form.Invoke(d);
            }
            return _form;
        }

        #endregion

    }
}
