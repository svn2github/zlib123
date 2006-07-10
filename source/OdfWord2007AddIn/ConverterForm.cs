/* 
 * Created by jgoffinet.
 * Date: 19/06/2006 at 14:08
 * 
 * This software is licensed under BSD-like license:
 * 
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
 *     * Neither the name of the University of California, Berkeley nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
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
using System.IO;
using System.Threading;
using CleverAge.OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.OdfWord2007Addin
{

    public partial class ConverterForm : Form
    {
        delegate void WorkCompleteCallback(Exception e);

        private string inputFile;
        private string outputFile;
        private bool computeSize;
        private int size;
        private Exception exception;
        private bool cancel;
        private bool converting;
        public ConverterForm(string inputFile, string outputFile)
        {
            InitializeComponent();
            this.inputFile = inputFile;
            this.outputFile = outputFile;
           // DoConvert();
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

        public bool Canceled {
            get {
                return cancel;
            }
        }
        private void DoConvert()
        {
            try {
                Converter converter = new Converter();
                converter.AddMessageListener(new Converter.MessageListener(MessageInterceptor));
                this.computeSize = true;
                converter.OdfToOoxComputeSize(this.inputFile);
                this.progressBar1.Maximum = this.size;
                this.computeSize = false;
                converter.OdfToOox(this.inputFile, this.outputFile);
                WorkComplete(null);
            } catch (Exception e) {
                WorkComplete(e);
            }
        }

        private void MessageInterceptor(object sender, EventArgs e)
        {
            if (this.computeSize)
            {
                this.size++;
            }
            else
            {
                this.progressBar1.Increment(1);
            }
            Application.DoEvents();
            if (cancel) {
                // As we need to leave converter.OdfToOox, throw an exception
                throw new CancelledException();
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

        private void cancelButton_Click(object sender, EventArgs e) {
            cancel = true;
        }

        private void ConverterForm_Load(object sender, EventArgs e) {
            FileInfo file = new FileInfo(inputFile);
            this.Text = "Converting " + file.Name + "...";
        }

        private void ConverterForm_Activated(object sender, EventArgs e) {
            // Launch convertion
            if (!converting) {
                converting = true;
                DoConvert();
                converting = false;
            }
        }
    }
}
