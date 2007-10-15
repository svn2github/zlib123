using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OdfConverterHost {
    public partial class Form1 : Form {





        public Form1() {
            InitializeComponent();
        }

        public void AjouterMessage(string message) {
            listBox1.Items.Add(message);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) {
            Application.Exit();
        }

        private void Form1_VisibleChanged(object sender, EventArgs e) {
            if (this.WindowState == FormWindowState.Minimized) {
                this.Visible = false;
            }
        }

        private void showLogToolStripMenuItem_Click(object sender, EventArgs e) {
            this.WindowState = FormWindowState.Normal;
            this.Visible = true;
        }

        private void Form1_Load(object sender, EventArgs e) {
            Win32.OutputDebugString("Setting Visible to False\n");
            this.Text = "Thread ID = " + Win32.GetCurrentThreadId().ToString();
        }

        private void btnConvertir_Click(object sender, EventArgs e) {
            try {
                AjouterMessage("Instanciation convertisseur");
                Converter converter = new Converter();
                AjouterMessage("Début conversion");
                converter.OdtToDocx(txtFichier.Text, @"C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\Gudule.docx", true, 1033, 0);
                AjouterMessage("Fin conversion");
            } catch (Exception ex) {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }
        }


    }
}