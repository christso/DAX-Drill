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
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblNamespace = new System.Windows.Forms.Label();
            this.cbNamespace = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // txtXmlText
            // 
            this.txtXmlText.Location = new System.Drawing.Point(12, 37);
            this.txtXmlText.Multiline = true;
            this.txtXmlText.Name = "txtXmlText";
            this.txtXmlText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtXmlText.Size = new System.Drawing.Size(566, 89);
            this.txtXmlText.TabIndex = 0;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(503, 132);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 2;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.BtnOkClick);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(407, 132);
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
            this.lblNamespace.Location = new System.Drawing.Point(13, 13);
            this.lblNamespace.Name = "lblNamespace";
            this.lblNamespace.Size = new System.Drawing.Size(64, 13);
            this.lblNamespace.TabIndex = 4;
            this.lblNamespace.Text = "Namespace";
            // 
            // cbNamespace
            // 
            this.cbNamespace.FormattingEnabled = true;
            this.cbNamespace.Location = new System.Drawing.Point(83, 10);
            this.cbNamespace.Name = "cbNamespace";
            this.cbNamespace.Size = new System.Drawing.Size(403, 21);
            this.cbNamespace.TabIndex = 5;
            // 
            // XmlEditForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(591, 162);
            this.Controls.Add(this.cbNamespace);
            this.Controls.Add(this.lblNamespace);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.txtXmlText);
            this.Name = "XmlEditForm";
            this.Text = "DAX Drill XML Configuration Editor";
            this.TopMost = true;
            this.Resize += new System.EventHandler(this.FormResizer);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TextBox txtXmlText;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblNamespace;
        private System.Windows.Forms.ComboBox cbNamespace;
    }
}