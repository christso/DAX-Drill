using System;
using System.Drawing;
using System.Windows.Forms;

namespace DaxDrill.UI
{
    /// <summary>
    /// Description of ErrForm.
    /// </summary>
    public partial class MsgForm : Form
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

        public static void ShowMessage(Exception ex, string formTitle = AppName)
        {
            ShowMessage(ex.Message, ex.ToString(), formTitle);
        }

        //Shows the static version of the form
        public static MsgForm ShowForm()
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

        private static MsgForm _form = new MsgForm();

        private delegate MsgForm GetFormCallBack();
        //Returns static version of the form
        public static MsgForm GetStatic()
        {

            //Reinstantiate the form
            if (_form.IsDisposed)
            {
                _form = new MsgForm();
            }

            if (_form.InvokeRequired)
            {
                var d = new GetFormCallBack(GetStatic);
                return (MsgForm)_form.Invoke(d);
            }
            return _form;
        }

        #endregion

        public MsgForm()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();

            //
            // Add constructor code after the InitializeComponent() call.
            //
        }

        #region Form Events

        void BtnOkClick(object sender, EventArgs e)
        {
            Hide();
        }

        #endregion
    }
}
