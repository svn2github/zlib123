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
using Microsoft.Win32;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    ///     Module ID:      Optional Dialog Box
    ///     Description:    Odf Converter 
    ///     Author:         Sateesh
    ///     Create Date:    2008-05-13
    /// </summary>
    public partial class ConfigForm : Form
    {
        string fidelityMsgValue=string.Empty;

        public ConfigForm()
        {

            Application.EnableVisualStyles();
            InitializeComponent();
            try
            {
                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "POWERPNT")
                {
                        string languageVal = Microsoft.Win32.Registry
                    .GetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", null) as string;

                        if (languageVal == null)
                        {
                            languageVal = Microsoft.Win32.Registry
                            .GetValue(@"HKEY_LOCAL_MACHINE\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", null) as string;

                            if (languageVal == "true")
                            {
                                chkbxIsErrorIgnored.Checked = true;
                                Microsoft.Win32.Registry
                       .SetValue(@"HKEY_LOCAL_MACHINE\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "true");
                            }
                            else if (languageVal == "false")
                            {
                                chkbxIsErrorIgnored.Checked = false;
                                Microsoft.Win32.Registry
                      .SetValue(@"HKEY_LOCAL_MACHINE\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "false");
                            }
                        }
                        else if (languageVal != null)
                        {

                            if (languageVal == "true")
                            {
                                chkbxIsErrorIgnored.Checked = true;
                                Microsoft.Win32.Registry
                       .SetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "true");
                            }
                            else if (languageVal == "false")
                            {
                                chkbxIsErrorIgnored.Checked = false;
                                Microsoft.Win32.Registry
                      .SetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "false");
                            }
                        }
                }
                else if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "EXCEL")
                {
                    string languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        languageVal = Microsoft.Win32.Registry
                        .GetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", null) as string;

                        if (languageVal == "true")
                        {
                            chkbxIsErrorIgnored.Checked = true;
                            Microsoft.Win32.Registry
                   .SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "true");
                        }
                        else
                        {
                            chkbxIsErrorIgnored.Checked = false;
                            Microsoft.Win32.Registry
                  .SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "false");
                        }
                    }
                    else if (languageVal != null)
                    {
                        if (languageVal == "true")
                        {
                            chkbxIsErrorIgnored.Checked = true;
                            Microsoft.Win32.Registry
                   .SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "true");
                        }
                        else
                        {
                            chkbxIsErrorIgnored.Checked = false;
                            Microsoft.Win32.Registry
                  .SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "false");
                        }
                    }
                }
                else if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "WINWORD")
                {
                    string languageVal = Microsoft.Win32.Registry
               .GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        languageVal = Microsoft.Win32.Registry
                        .GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Clever Age\Odf Add-in for Word", "fidelityValue", null) as string;

                        if (languageVal == "true")
                        {
                            chkbxIsErrorIgnored.Checked = true;
                            Microsoft.Win32.Registry
                   .SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "true");
                        }
                        else
                        {
                            chkbxIsErrorIgnored.Checked = false;
                            Microsoft.Win32.Registry
                  .SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "false");
                        }
                    }
                    else if (languageVal != null)
                    {
                        if (languageVal == "true")
                        {
                            chkbxIsErrorIgnored.Checked = true;
                            Microsoft.Win32.Registry
                   .SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "true");
                        }
                        else
                        {
                            chkbxIsErrorIgnored.Checked = false;
                            Microsoft.Win32.Registry
                  .SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "false");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {
            
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (chkbxIsErrorIgnored.Checked)
            {
                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "POWERPNT")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "true");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "true");
                    }
                }

                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "EXCEL")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "true");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "true");
                    }
                }

                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "WINWORD")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "true");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "true");
                    }
                }
            }
            else
            {
                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "POWERPNT")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "false");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Presentation", "fidelityValue", "false");
                    }
                }

                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "EXCEL")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "false");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "fidelityValue", "false");
                    }
                }

                if (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "WINWORD")
                {
                    string languageVal = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", null) as string;

                    if (languageVal == null)
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "false");
                    }
                    else
                    {
                        Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "fidelityValue", "false");
                    }
                }


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