namespace OdfConverter.OdfConverterLib
{
    partial class InfoBox
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InfoBox));
            this.label = new System.Windows.Forms.Label();
            this.OK = new System.Windows.Forms.Button();
            this.Details = new System.Windows.Forms.Button();
            this.grpDetails = new System.Windows.Forms.GroupBox();
            this.txtDetails = new System.Windows.Forms.TextBox();
            this.chkbxIsErrorIgnored = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpDetails.SuspendLayout();
            this.SuspendLayout();
            // 
            // label
            // 
            this.label.BackColor = System.Drawing.SystemColors.Control;
            resources.ApplyResources(this.label, "label");
            this.label.Name = "label";
            // 
            // OK
            // 
            this.OK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.OK, "OK");
            this.OK.Name = "OK";
            this.OK.UseVisualStyleBackColor = true;
            // 
            // Details
            // 
            resources.ApplyResources(this.Details, "Details");
            this.Details.Name = "Details";
            this.Details.UseVisualStyleBackColor = true;
            this.Details.Click += new System.EventHandler(this.Details_Click);
            // 
            // grpDetails
            // 
            this.grpDetails.Controls.Add(this.label1);
            this.grpDetails.Controls.Add(this.txtDetails);
            resources.ApplyResources(this.grpDetails, "grpDetails");
            this.grpDetails.Name = "grpDetails";
            this.grpDetails.TabStop = false;
            // 
            // txtDetails
            // 
            resources.ApplyResources(this.txtDetails, "txtDetails");
            this.txtDetails.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDetails.Name = "txtDetails";
            this.txtDetails.ReadOnly = true;
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // chkbxIsErrorIgnored
            // 
            resources.ApplyResources(this.chkbxIsErrorIgnored, "chkbxIsErrorIgnored");
            this.chkbxIsErrorIgnored.Name = "chkbxIsErrorIgnored";
            this.chkbxIsErrorIgnored.UseVisualStyleBackColor = true;
            this.chkbxIsErrorIgnored.CheckedChanged += new System.EventHandler(this.chkbxIsErrorIgnored_CheckedChanged);
            // 
            // InfoBox
            // 
            this.AcceptButton = this.OK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.chkbxIsErrorIgnored);
            this.Controls.Add(this.grpDetails);
            this.Controls.Add(this.OK);
            this.Controls.Add(this.Details);
            this.Controls.Add(this.label);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InfoBox";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Load += new System.EventHandler(this.InfoBox_Load);
            this.grpDetails.ResumeLayout(false);
            this.grpDetails.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Button Details;
        private System.Windows.Forms.GroupBox grpDetails;
        private System.Windows.Forms.TextBox txtDetails;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkbxIsErrorIgnored;
    }
}