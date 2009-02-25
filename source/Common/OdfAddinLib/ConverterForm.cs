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
using OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    public partial class ConverterForm : Form
    {
        delegate void WorkCompleteCallback(Exception e);

        private AbstractConverter _converter;
        //private string _inputFile;
        //private string _outputFile;
        private ResourceManager _manager;
        private ConversionOptions _options;
        private bool _computeSize;
        private int _size;
        private Exception _exception;
        private bool _cancel;
        private bool _converting;
        private List<string> _lostElements;

        public ConverterForm(AbstractConverter converter, string inputFile, string outputFile, ResourceManager manager, ConversionOptions options)
        {
            InitializeComponent();

            this._converter = converter;
            this._manager = manager;
            this._options = options;
            this._lostElements = new List<string>();
            
            this.lblMessage.Text = manager.GetString("ProgressBarLoadLabel");
            this.lblMessage.Visible = true;

            FileInfo file = new FileInfo(options.InputFullNameOriginal);
            this.Text = _manager.GetString("ConversionFormTitle").Replace("%1", file.Name);

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
            get { return this._exception; }
        }

        public bool Canceled
        {
            get { return _cancel; }
        }

        public bool HasLostElements
        {
            get { return _lostElements.Count > 0; }
        }

        public string[] LostElements
        {
            get { return (string[])_lostElements.ToArray(); }
        }

        private void DoConvert()
        {
            try
            {
                _converter.RemoveMessageListeners();
                _converter.AddProgressMessageListener(new AbstractConverter.MessageListener(ProgressMessageInterceptor));
                _converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
                _converter.DirectTransform = this._options.IsDirectTransform;

                if (UriLoader.IsRemote(this._options.InputFullName))
                {
                    this._options.InputFullNameOriginal = this._options.InputFullName;
                    this._options.InputFullName = UriLoader.DownloadFile(this._options.InputFullName);
                    this._options.InputBaseFolder = Path.GetDirectoryName(this._options.InputFullName);    
                }
                
                this._computeSize = true;
                this._converter.ComputeSize(this._options.InputFullName);
                this.progressBar1.Maximum = this._size;
                this._computeSize = false;
                _converter.Transform(this._options.InputFullName, this._options.OutputFullName, this._options);
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
            if (this._computeSize)
            {
                this._size++;
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
                    this.lblMessage.Text = _manager.GetString("ProgressBarExitLabel");
                }
            }
            Application.DoEvents();
            if (_cancel)
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
                messageValue = _manager.GetString(messageKey.Substring(0, index));

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
                messageValue = _manager.GetString(messageKey);
            }

            if (messageValue != null && !_lostElements.Contains(messageValue))
            {
                _lostElements.Add(messageValue);
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
                    this._exception = e;
                    DialogResult = DialogResult.Abort;
                }
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            _cancel = true;
        }

        private void ConverterForm_Activated(object sender, EventArgs e)
        {
            // Launch conversion
            if (!_converting)
            {
                _converting = true;
                Application.DoEvents();
                DoConvert();
                _converting = false;
            }
        }
    }
}
