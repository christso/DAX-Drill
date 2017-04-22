using System;

namespace DaxDrill.UI
{
    partial class AboutBox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Label1 = new System.Windows.Forms.Label();
            this.txtAddInPath = new System.Windows.Forms.TextBox();
            this.lblAddInPath = new System.Windows.Forms.Label();
            this.lblAuthorName = new System.Windows.Forms.Label();
            this.lblVersionNum = new System.Windows.Forms.Label();
            this.lblAuthor = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lblWebsite = new System.Windows.Forms.Label();
            this.lblWebsiteAddr = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label1.Location = new System.Drawing.Point(12, 9);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(157, 13);
            this.Label1.TabIndex = 20;
            this.Label1.Text = "DAX OLAP Drill-Through Add-In";
            // 
            // txtAddInPath
            // 
            this.txtAddInPath.Location = new System.Drawing.Point(12, 112);
            this.txtAddInPath.Name = "txtAddInPath";
            this.txtAddInPath.Size = new System.Drawing.Size(171, 20);
            this.txtAddInPath.TabIndex = 19;
            this.txtAddInPath.Text = "Unknown";
            // 
            // lblAddInPath
            // 
            this.lblAddInPath.AutoSize = true;
            this.lblAddInPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAddInPath.Location = new System.Drawing.Point(12, 95);
            this.lblAddInPath.Name = "lblAddInPath";
            this.lblAddInPath.Size = new System.Drawing.Size(85, 13);
            this.lblAddInPath.TabIndex = 18;
            this.lblAddInPath.Text = "Add-In File Path:";
            // 
            // lblAuthorName
            // 
            this.lblAuthorName.AutoSize = true;
            this.lblAuthorName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAuthorName.Location = new System.Drawing.Point(65, 73);
            this.lblAuthorName.Name = "lblAuthorName";
            this.lblAuthorName.Size = new System.Drawing.Size(81, 13);
            this.lblAuthorName.TabIndex = 17;
            this.lblAuthorName.Text = "Christopher Tso";
            // 
            // lblVersionNum
            // 
            this.lblVersionNum.AutoSize = true;
            this.lblVersionNum.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblVersionNum.Location = new System.Drawing.Point(65, 53);
            this.lblVersionNum.Name = "lblVersionNum";
            this.lblVersionNum.Size = new System.Drawing.Size(31, 13);
            this.lblVersionNum.TabIndex = 16;
            this.lblVersionNum.Text = "1.0.0";
            // 
            // lblAuthor
            // 
            this.lblAuthor.AutoSize = true;
            this.lblAuthor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAuthor.Location = new System.Drawing.Point(12, 73);
            this.lblAuthor.Name = "lblAuthor";
            this.lblAuthor.Size = new System.Drawing.Size(41, 13);
            this.lblAuthor.TabIndex = 15;
            this.lblAuthor.Text = "Author:";
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.Location = new System.Drawing.Point(12, 53);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(45, 13);
            this.lblVersion.TabIndex = 14;
            this.lblVersion.Text = "Version:";
            // 
            // lblWebsite
            // 
            this.lblWebsite.AutoSize = true;
            this.lblWebsite.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWebsite.Location = new System.Drawing.Point(12, 33);
            this.lblWebsite.Name = "lblWebsite";
            this.lblWebsite.Size = new System.Drawing.Size(49, 13);
            this.lblWebsite.TabIndex = 13;
            this.lblWebsite.Text = "Website:";
            // 
            // lblWebsiteAddr
            // 
            this.lblWebsiteAddr.AutoSize = true;
            this.lblWebsiteAddr.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblWebsiteAddr.Location = new System.Drawing.Point(65, 33);
            this.lblWebsiteAddr.Name = "lblWebsiteAddr";
            this.lblWebsiteAddr.Size = new System.Drawing.Size(183, 13);
            this.lblWebsiteAddr.TabIndex = 12;
            this.lblWebsiteAddr.Text = "https://github.com/christso/DAX-Drill";
            this.lblWebsiteAddr.Click += new System.EventHandler(this.lblWebsiteAddr_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(191, 110);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "&OK";
            // 
            // AboutBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(283, 145);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.txtAddInPath);
            this.Controls.Add(this.lblAddInPath);
            this.Controls.Add(this.lblAuthorName);
            this.Controls.Add(this.lblVersionNum);
            this.Controls.Add(this.lblAuthor);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.lblWebsite);
            this.Controls.Add(this.lblWebsiteAddr);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "AboutBox";
            this.Text = "About DAX Drill";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.AboutBox_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        #endregion

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox txtAddInPath;
        internal System.Windows.Forms.Label lblAddInPath;
        internal System.Windows.Forms.Label lblAuthorName;
        internal System.Windows.Forms.Label lblVersionNum;
        internal System.Windows.Forms.Label lblAuthor;
        internal System.Windows.Forms.Label lblVersion;
        internal System.Windows.Forms.Label lblWebsite;
        internal System.Windows.Forms.Label lblWebsiteAddr;
        internal System.Windows.Forms.Button btnOK;
    }
}