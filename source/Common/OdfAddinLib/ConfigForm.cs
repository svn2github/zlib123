/*
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    ///     Module ID:      
    ///     Description:    OdfConverter Customization
    ///     Author:         Sateesh
    ///     Create Date:    2008-05-13
    /// </summary>
    public partial class ConfigForm : Form
    {
        private ConfigManager configMan;

        public ConfigForm()
        {

            Application.EnableVisualStyles();
            
            InitializeComponent();

            try
            {
                configMan = new ConfigManager(System.IO.Path.GetDirectoryName(typeof(ConverterForm).Assembly.Location) + @"\conf\config.xml");
                configMan.LoadConfig();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            chkbxIsErrorIgnored.Checked = configMan.IsErrorIgnored;
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Save Config
            configMan.IsErrorIgnored = chkbxIsErrorIgnored.Checked;

            try
            {
                configMan.SaveConfig();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ConfigForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }
    }
}