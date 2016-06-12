namespace DG2NTT.DaxDrill.UI
{
    partial class XmlEditForm
    {
        /// <summary>
        /// Designer variable used to keep track of non-visual components.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Disposes resources used by the form.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// This method is required for Windows Forms designer support.
        /// Do not change the method contents inside the source code editor. The Forms designer might
        /// not be able to load this method if it was changed manually.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtXmlText = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblNamespace = new System.Windows.Forms.Label();
            this.cbNamespace = new System.Windows.Forms.ComboBox();
            this.cbWorkbooks = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.lblXpath = new System.Windows.Forms.Label();
            this.txtXpath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtXmlText
            // 
            this.txtXmlText.Location = new System.Drawing.Point(12, 66);
            this.txtXmlText.Multiline = true;
            this.txtXmlText.Name = "txtXmlText";
            this.txtXmlText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtXmlText.Size = new System.Drawing.Size(566, 152);
            this.txtXmlText.TabIndex = 0;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(503, 224);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSaveClick);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(422, 224);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblNamespace
            // 
            this.lblNamespace.AutoSize = true;
            this.lblNamespace.Location = new System.Drawing.Point(302, 15);
            this.lblNamespace.Name = "lblNamespace";
            this.lblNamespace.Size = new System.Drawing.Size(64, 13);
            this.lblNamespace.TabIndex = 4;
            this.lblNamespace.Text = "Namespace";
            // 
            // cbNamespace
            // 
            this.cbNamespace.FormattingEnabled = true;
            this.cbNamespace.Location = new System.Drawing.Point(372, 12);
            this.cbNamespace.Name = "cbNamespace";
            this.cbNamespace.Size = new System.Drawing.Size(207, 21);
            this.cbNamespace.TabIndex = 5;
            // 
            // cbWorkbooks
            // 
            this.cbWorkbooks.FormattingEnabled = true;
            this.cbWorkbooks.Location = new System.Drawing.Point(84, 12);
            this.cbWorkbooks.Name = "cbWorkbooks";
            this.cbWorkbooks.Size = new System.Drawing.Size(194, 21);
            this.cbWorkbooks.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Workbook";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(341, 224);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 23);
            this.btnRefresh.TabIndex = 8;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // lblXpath
            // 
            this.lblXpath.AutoSize = true;
            this.lblXpath.Location = new System.Drawing.Point(330, 42);
            this.lblXpath.Name = "lblXpath";
            this.lblXpath.Size = new System.Drawing.Size(36, 13);
            this.lblXpath.TabIndex = 9;
            this.lblXpath.Text = "XPath";
            // 
            // txtXpath
            // 
            this.txtXpath.Location = new System.Drawing.Point(372, 40);
            this.txtXpath.Name = "txtXpath";
            this.txtXpath.Size = new System.Drawing.Size(206, 20);
            this.txtXpath.TabIndex = 10;
            // 
            // XmlEditForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 254);
            this.Controls.Add(this.txtXpath);
            this.Controls.Add(this.lblXpath);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.cbWorkbooks);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbNamespace);
            this.Controls.Add(this.lblNamespace);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtXmlText);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "XmlEditForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "DAX Drill XML Editor";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtXmlText;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblNamespace;
        private System.Windows.Forms.ComboBox cbNamespace;
        private System.Windows.Forms.ComboBox cbWorkbooks;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Label lblXpath;
        private System.Windows.Forms.TextBox txtXpath;
    }
}