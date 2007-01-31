namespace CleverAge.OdfConverter.OdfWord2003Addin
{
    partial class FrmPrerequisites
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPrerequisites));
            this.panelTitle = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblTopLine = new System.Windows.Forms.Label();
            this.lblBottomLine = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.linkAddRemove = new System.Windows.Forms.LinkLabel();
            this.linkOfficePIA = new System.Windows.Forms.LinkLabel();
            this.txtPart1 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panelTitle
            // 
            this.panelTitle.BackColor = System.Drawing.Color.White;
            this.panelTitle.Controls.Add(this.lblTitle);
            this.panelTitle.Controls.Add(this.pictureBox1);
            this.panelTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTitle.Location = new System.Drawing.Point(0, 0);
            this.panelTitle.Name = "panelTitle";
            this.panelTitle.Size = new System.Drawing.Size(497, 69);
            this.panelTitle.TabIndex = 0;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(8, 12);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(332, 20);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "ODF Add-in for Word 2003 Prerequisites";
            // 
            // lblTopLine
            // 
            this.lblTopLine.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTopLine.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTopLine.Location = new System.Drawing.Point(0, 69);
            this.lblTopLine.Name = "lblTopLine";
            this.lblTopLine.Size = new System.Drawing.Size(497, 2);
            this.lblTopLine.TabIndex = 2;
            // 
            // lblBottomLine
            // 
            this.lblBottomLine.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblBottomLine.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBottomLine.Location = new System.Drawing.Point(0, 336);
            this.lblBottomLine.Name = "lblBottomLine";
            this.lblBottomLine.Size = new System.Drawing.Size(500, 2);
            this.lblBottomLine.TabIndex = 5;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(400, 348);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(88, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // btnBack
            // 
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBack.Enabled = false;
            this.btnBack.Location = new System.Drawing.Point(304, 348);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(88, 23);
            this.btnBack.TabIndex = 6;
            this.btnBack.Text = "< &Back";
            this.btnBack.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(208, 348);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // linkAddRemove
            // 
            this.linkAddRemove.AutoSize = true;
            this.linkAddRemove.Location = new System.Drawing.Point(52, 182);
            this.linkAddRemove.Name = "linkAddRemove";
            this.linkAddRemove.Size = new System.Drawing.Size(314, 13);
            this.linkAddRemove.TabIndex = 8;
            this.linkAddRemove.TabStop = true;
            this.linkAddRemove.Text = "Launch Add/Remove program applet to update Office installation";
            this.linkAddRemove.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkAddRemove_LinkClicked);
            // 
            // linkOfficePIA
            // 
            this.linkOfficePIA.AutoSize = true;
            this.linkOfficePIA.Location = new System.Drawing.Point(52, 205);
            this.linkOfficePIA.Name = "linkOfficePIA";
            this.linkOfficePIA.Size = new System.Drawing.Size(390, 13);
            this.linkOfficePIA.TabIndex = 9;
            this.linkOfficePIA.TabStop = true;
            this.linkOfficePIA.Text = "Download Office 2003 Primary Interop Assemblies from http://www.microsoft.com";
            this.linkOfficePIA.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkOfficePIA_LinkClicked);
            // 
            // txtPart1
            // 
            this.txtPart1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtPart1.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.txtPart1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPart1.Location = new System.Drawing.Point(52, 90);
            this.txtPart1.Multiline = true;
            this.txtPart1.Name = "txtPart1";
            this.txtPart1.ReadOnly = true;
            this.txtPart1.Size = new System.Drawing.Size(433, 43);
            this.txtPart1.TabIndex = 7;
            this.txtPart1.Text = "Office 2003 Primary Interop Assemblies are not installed on this computer.\r\nYou n" +
                "eed to install them before installing this Add-In for Word 2003.";
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.textBox1.Location = new System.Drawing.Point(52, 139);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(395, 34);
            this.textBox1.TabIndex = 10;
            this.textBox1.Text = "They can be added through Office 2003 Setup (select \".Net programmability\" for Wo" +
                "rd application) or downloaded from Microsoft\'s Web site.\r\n";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::CleverAge.OdfConverter.OdfWord2003Addin.Properties.Resources.Error;
            this.pictureBox2.Location = new System.Drawing.Point(14, 87);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(32, 32);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox2.TabIndex = 11;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = global::CleverAge.OdfConverter.OdfWord2003Addin.Properties.Resources.ImageInstaller;
            this.pictureBox1.Location = new System.Drawing.Point(413, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(84, 63);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // FrmPrerequisites
            // 
            this.AcceptButton = this.btnClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(497, 382);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.linkAddRemove);
            this.Controls.Add(this.linkOfficePIA);
            this.Controls.Add(this.txtPart1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblBottomLine);
            this.Controls.Add(this.lblTopLine);
            this.Controls.Add(this.panelTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmPrerequisites";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "MyTitle";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmPrerequisites_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panelTitle.ResumeLayout(false);
            this.panelTitle.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panelTitle;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblTopLine;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblBottomLine;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.LinkLabel linkAddRemove;
        private System.Windows.Forms.LinkLabel linkOfficePIA;
        private System.Windows.Forms.TextBox txtPart1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}

