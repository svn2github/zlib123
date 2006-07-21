using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Resources;
using System.Collections;

namespace CleverAge.OdfConverter.OdfWordAddinLib
{
    public partial class InfoBox : Form
    {
        private ResourceManager manager;
        private bool showDetails;

        public InfoBox(ArrayList elements, ResourceManager manager)
        {
            InitializeComponent();
            this.manager = manager;
            showDetails = false;
            foreach (string element in elements)
            {
                textBox1.Text += element + "\r\n";
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void InfoBox_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            showDetails = !showDetails;
            textBox1.Visible = showDetails;
        }

        private void OK_Click(object sender, EventArgs e)
        {

        }
    }
}