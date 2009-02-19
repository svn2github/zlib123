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
using System.IO;
using System.Resources;
using System.Windows.Forms;
using System.Net;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    public partial class ConverterForm : Form
    {
        delegate void WorkCompleteCallback(Exception e);

        private AbstractConverter converter;
        private string inputFile;
        private string outputFile;
        private ResourceManager manager;
        private ConversionOptions options;
        private bool computeSize;
        private int size;
        private Exception exception;
        private bool cancel;
        private bool converting;
        private List<string> lostElements;

        public ConverterForm(AbstractConverter converter, string inputFile, string outputFile, ResourceManager manager, ConversionOptions options)
        {
            InitializeComponent();
            this.converter = converter;
            this.inputFile = inputFile;
            this.outputFile = outputFile;
            this.manager = manager;
            this.options = options;
            lostElements = new List<string>();
            //this code is for displaying the label in progress bar               
            //Code change 1 of 2
            this.lblMessage.Text = manager.GetString("ProgressBarLoadLabel");
            this.lblMessage.Visible = true;

            if (this.Parent == null)
            {
                // started in stand-alone mode (e.g. via context menu)
                this.StartPosition = FormStartPosition.CenterScreen;
            }
        }

        public class CancelledException : Exception
        {
        }

        public Exception Exception
        {
            get
            {
                return this.exception;
            }
        }

        public bool Canceled
        {
            get
            {
                return cancel;
            }
        }

        public bool HasLostElements
        {
            get
            {
                return lostElements.Count > 0;
            }
        }

        public string[] LostElements
        {
            get
            {
                return (string[])lostElements.ToArray();
            }
        }

        private void DoConvert()
        {
            try
            {
                converter.RemoveMessageListeners();
                converter.AddProgressMessageListener(new AbstractConverter.MessageListener(ProgressMessageInterceptor));
                converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
                converter.DirectTransform = this.options.IsDirectTransform;
                this.computeSize = true;
                converter.ComputeSize(this.inputFile);
                this.progressBar1.Maximum = this.size;
                this.computeSize = false;
                //Code Added on 19 Feb 2009 by Achougle@Xandros.com
                this.options.InputFullName = this.inputFile;
                this.options.InputFullNameOriginal = this.inputFile;
                //Code Added on 19 Feb 2009 by Achougle@Xandros.com
                converter.Transform(this.inputFile, this.outputFile, this.options);
                WorkComplete(null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                WorkComplete(ex);
            }
        }

        private void ProgressMessageInterceptor(object sender, EventArgs e)
        {
            if (this.computeSize)
            {
                this.size++;
            }
            else
            {
                this.progressBar1.Increment(1);

                //this code is for displaying the label in progress bar                         
                //Code change 2 of 2                
                if (this.progressBar1.Value == 1)
                //when progress bar value is 1 then label is made in visible
                {
                    this.lblMessage.Visible = false;
                }
                if (this.progressBar1.Value >= this.progressBar1.Maximum)
                //when progress bar crosses maximum value, another message will be displayed 
                {
                    this.lblMessage.Visible = true;
                    this.lblMessage.Text = manager.GetString("ProgressBarExitLabel");
                }
            }
            Application.DoEvents();
            if (cancel)
            {
                // As we need to leave converter.OdfToOox, throw an exception
                throw new CancelledException();
            }
        }

        private void FeedbackMessageInterceptor(object sender, EventArgs e)
        {
            string messageKey = ((OdfEventArgs)e).Message;
            string messageValue = null;

            int index = messageKey.IndexOf('%');
            // parameters substitution
            if (index > 0)
            {
                string[] param = messageKey.Substring(index + 1).Split(new char[] { '%' });
                messageValue = manager.GetString(messageKey.Substring(0, index));

                if (messageValue != null)
                {
                    for (int i = 0; i < param.Length; i++)
                    {
                        messageValue = messageValue.Replace("%" + (i + 1), param[i]);
                    }
                }
            }
            else
            {
                messageValue = manager.GetString(messageKey);
            }

            if (messageValue != null && !lostElements.Contains(messageValue))
            {
                lostElements.Add(messageValue);
            }
        }

        private void WorkComplete(Exception e)
        {
            if (e == null)
            {
                DialogResult = DialogResult.OK;
            }
            else
            {
                if (e is CancelledException)
                {
                    DialogResult = DialogResult.Cancel;
                }
                else
                {
                    this.exception = e;
                    DialogResult = DialogResult.Abort;
                }
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            cancel = true;
        }

        private void ConverterForm_Load(object sender, EventArgs e)
        {
            //Code Added on 19 Feb 2009 by Achougle@Xandros.com
            String fileurl = inputFile;
            //Chcking whether the inputFile is url of ODF document
            if (fileurl.Contains("http:") || fileurl.Contains("ftp:"))
            {
                //Extracting the Filename from the url.
                String filename = fileurl.Substring(fileurl.LastIndexOf("/"), (fileurl.Length - fileurl.LastIndexOf("/")));
                filename = filename.Remove(0, 1);
                string tempPath = Path.GetTempPath().ToString();
                String str = tempPath + filename;
                /*Using WebClient class provides common methods for sending data to or receiving data from
                any local, intranet, or Internet resource identified by a URI.*/
                WebClient wc = new WebClient();
                try
                {
                    //Downloading the ODF document to the user's local temp directory.
                    wc.DownloadFile(inputFile, str);
                }
                catch (Exception e1)
                {
                    this.exception = e1;
                }
                //Assigning the local temp path filename to the inputFile.
                inputFile = str;
            }
            //Code Added on 19 Feb 2009 by Achougle@Xandros.com
            FileInfo file = new FileInfo(inputFile);
            this.Text = manager.GetString("ConversionFormTitle").Replace("%1", file.Name);
        }

        private void ConverterForm_Activated(object sender, EventArgs e)
        {
            // Launch conversion
            if (!converting)
            {
                converting = true;
                Application.DoEvents();
                DoConvert();
                converting = false;
            }
        }
    }
}
