using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DaxDrill.Helpers;
using DaxDrill.ExcelHelpers;

namespace DaxDrill.UI
{
    public partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();
        }

        private static AboutBox oForm = new AboutBox();

        //Returns static version of the form
        public static AboutBox GetForm()
        {
            return oForm;
        }

        //Shows the static version of the form
        public static AboutBox ShowForm()
        {
            if (oForm.IsDisposed)
            {
                oForm = new AboutBox();
            }

            //Show Form
            oForm.ShowDialog();
            oForm.TopMost = true;
            //keep form on top
            return (oForm);
        }

        private void AboutBox_Load(object sender, EventArgs e)
        {
            // Set the title of the form.
            this.Text = "About DAX Drill";

            // Initialize all of the text displayed on the About Box.
            this.lblVersionNum.Text = string.Format("{0}", AssemblyHelper.AssemblyVersion);

            this.txtAddInPath.Text = ExcelHelper.AddInPath;
        }

        private void lblWebsiteAddr_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(lblWebsiteAddr.Text);
        }
    }
}
