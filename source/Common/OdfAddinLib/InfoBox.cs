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
using System.Resources;
using System.Collections;
using Microsoft.Win32;

using System.Runtime.InteropServices;
namespace CleverAge.OdfConverter.OdfConverterLib
{
    public partial class InfoBox : Form
    {
        private ResourceManager manager;
        /// <summary>
        /// Are details shown
        /// </summary>
        private bool showDetails;
        /// <summary>
        /// Client size of dialog box in "no details" mode
        /// </summary>
        private Size smallSize = new Size(387, 65);
        /// <summary>
        /// Client size of dialog box in "show details" mode
        /// </summary>
        private Size largeSize;



        public InfoBox(string label, ArrayList details, ResourceManager manager)
        {
            InitializeComponent();
            this.manager = manager;
            this.label.Text = manager.GetString(label);
            StringBuilder bld = new StringBuilder();
            foreach (string detail in details)
            {
                string text = manager.GetString(detail);
                bld.Append(string.IsNullOrEmpty(text) ? detail : text);
                bld.Append("\r\n");
            }
            txtDetails.Text = bld.ToString();

            if (this.Parent == null)
            {
                // started in stand-alone mode (e.g. via context menu)
                this.StartPosition = FormStartPosition.CenterScreen;
            }
        }

        public InfoBox(string label, string details, ResourceManager manager)
        {
            InitializeComponent();
            this.manager = manager;
            this.label.Text = manager.GetString(label);
            string text = manager.GetString(details);
            txtDetails.Text = (string.IsNullOrEmpty(text) ? details : text);
        }

        private void InfoBox_Load(object sender, EventArgs e)
        {
            try {
                // Change the title
                string newTitle = manager.GetString("OdfConverterTitle");
                if (!string.IsNullOrEmpty(newTitle)) {
                    Text = newTitle;
                }
                // Store the offsets of buttons and groupbox
                Size proposedSize = new Size(label.Width, 3000); // No vertical constraint
                Size newSize = label.GetPreferredSize(proposedSize);
                int newHeight = newSize.Height;
                int offset = newHeight - label.Height;

                // Redim/move windows and controls
                label.Height = newHeight;
                smallSize.Height += offset;
                chkbxIsErrorIgnored.Top = label.Height ;
                OK.Top = chkbxIsErrorIgnored.Height+ chkbxIsErrorIgnored.Top ;
                Details.Top = chkbxIsErrorIgnored.Height + chkbxIsErrorIgnored.Top;

                // Test if everything fits
                int offsetRight = 0;
                int offsetBottom = 0;
                int leftMargin = label.Left;
                if (Details.Left < OK.Right) {
                    offsetRight = (OK.Right + leftMargin - Details.Left);
                    Details.Left += offsetRight;
                    grpDetails.Width += offsetRight;
                }
                if (Details.Right + leftMargin > smallSize.Width) {
                    offsetRight += Details.Right + leftMargin - smallSize.Width;
                    this.Width += offsetRight;
                    smallSize.Width += offsetRight;
                }
                if (Details.Bottom + leftMargin > smallSize.Height) {
                    offsetBottom = Details.Bottom + leftMargin - smallSize.Height;
                    this.Height += offsetBottom;
                    smallSize.Height += offsetBottom;
                }

                grpDetails.Top = Details.Top + Details.Height + 10;

                // Now compute the size needed for details text box and form with details are shown
                int marginBottom = ClientSize.Height - grpDetails.Bottom;

                label1.Text = txtDetails.Text + "_";
                proposedSize = new Size(label1.Width, 2000);
                newSize = label1.GetPreferredSize(proposedSize);
                newHeight = newSize.Height;
                if (newHeight > txtDetails.Height) {
                    // Will add scrollbars
                    txtDetails.ScrollBars = ScrollBars.Vertical;
                } else {
                    // No scrollbar needed
                    int offset2 = newHeight - txtDetails.Height;
                    txtDetails.Height = newHeight;
                    txtDetails.ScrollBars = ScrollBars.None;
                    grpDetails.Height += offset2;
                }
                largeSize = smallSize;
                largeSize.Height = grpDetails.Top + grpDetails.Height+ 10;
                // At loadtime : no details
                this.showDetails = false;
                this.ClientSize = smallSize;
                txtDetails.Visible = showDetails;
                grpDetails.Visible = showDetails;
            } catch {
                // No message no crash
            }
        }

        private void Details_Click(object sender, EventArgs e)
        {
            showDetails = !showDetails;

            txtDetails.Visible = showDetails;
            grpDetails.Visible = showDetails;
            if (showDetails)
            {
                this.ClientSize = largeSize;
                Details.Text = Details.Text.Replace("> > >", "< < <");
            }
            else
            {
                this.ClientSize = smallSize;
                Details.Text = Details.Text.Replace("< < <", "> > >");
            }
        }

        private void chkbxIsErrorIgnored_CheckedChanged(object sender, EventArgs e)
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
}
    }
}