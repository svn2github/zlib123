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
 */

using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Reflection;
using System.Collections;
using CleverAge.OdfConverter.OdfZipUtils;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    /// Core conversion methods 
    /// </summary>
    public abstract class AbstractConverter
    {
        private const string ODFToOOX_COMPUTE_SIZE_XSL = "odf2oox-compute-size.xsl";
        private const string OOXToODF_COMPUTE_SIZE_XSL = "oox2odf-compute-size.xsl";
      
        private bool isDirectTransform = true;
        private ArrayList skipedPostProcessors = null;
        private string externalResource = null;
        private bool packaging = true;
        
        protected AbstractConverter()
        {
            this.skipedPostProcessors = new ArrayList();
        }

        public bool DirectTransform
        {
            set { this.isDirectTransform = value; }
            get { return this.isDirectTransform; }
        }

        public ArrayList SkipedPostProcessors
        {
            set { this.skipedPostProcessors = value; }
        }

        public string ExternalResources
        {
            set { this.externalResource = value; }
            get { return this.externalResource; }
        }

        public bool Packaging
        {
            set { this.packaging = value; }
        }

        /// <summary>
        /// Specify a chain of post processors to be hooked to the xslt output for the direct conversion
        /// </summary>
        protected abstract string[] DirectPostProcessorsChain
        {
            get;
        }

        /// <summary>
        /// Specify a chain of post processors to be hooked to the xslt output for the reverse conversion
        /// </summary>
        protected abstract string[] ReversePostProcessorsChain
        {
            get;
        }

        /// <summary>
        /// Specify how to resolve the xsl documents uri's
        /// </summary>
        protected abstract XmlUrlResolver ResourceResolver
        {
            get;
        }

        /// <summary>
        /// Get the input xml document to the xsl transformation
        /// </summary>
        protected abstract XmlReader Source
        {
            get;
        }

        /// <summary>
        /// Get the input xsl document to the xsl transformation
        /// </summary>
        protected abstract XPathDocument XslDoc
        {
            get;
        }

        /// <summary>
        /// Get the xslt settings
        /// </summary>
        protected abstract XsltSettings XsltProcSettings
        {
            get;
        }

        public delegate void MessageListener(object sender, EventArgs e);

        private event MessageListener progressMessageIntercepted;
        private event MessageListener feedbackMessageIntercepted;

        public void AddProgressMessageListener(MessageListener listener)
        {
            progressMessageIntercepted += listener;
        }

        public void AddFeedbackMessageListener(MessageListener listener)
        {
            feedbackMessageIntercepted += listener;
        }

        public void ComputeSize(string inputFile)
        {
            Transform(inputFile, null);
        }

        /// <summary>
        /// bug #1644285 Zlib crashes on non-ascii file names.
        /// </summary>
        public void _Transform(string inputFile, string outputFile)
        {
            // Get the \Temp path
            string tempInputFile = Path.GetTempPath().ToString() + "odf-converter.input";
            string tempOutputFile = outputFile == null ? null : Path.GetTempPath().ToString() + "odf-converter.output";

            try
            {
                File.Copy(inputFile, tempInputFile, true);
                _Transform(tempInputFile, tempOutputFile);

                if (outputFile != null)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    File.Move(tempOutputFile, outputFile);
                }
            }
            finally
            {
                if (File.Exists(tempInputFile))
                {
                    File.Delete(tempInputFile);
                }
            }
        }

        public void Transform(string inputFile, string outputFile)
        {
            // this throws an exception in the the following cases:
            // - input file is not a valid file
            // - input file is an encrypted file
            CheckFile(inputFile);

            XmlReader source = null;
            XmlWriter writer = null;
            ZipResolver zipResolver = null;
           
            try
            {
                // create a xsl transformer
                XslCompiledTransform xslt = new XslCompiledTransform();
                
                // compile the stylesheet. 
                // Input stylesheet, xslt settings and uri resolver are retrieve from the implementation class.
                xslt.Load(this.XslDoc, this.XsltProcSettings, this.ResourceResolver);

                zipResolver = new ZipResolver(inputFile);
                XsltArgumentList parameters = new XsltArgumentList();
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                
                if (outputFile != null)
                {
                    parameters.AddParam("outputFile", "", outputFile);
                    XmlWriter finalWriter;
                    if (this.packaging)
                    {
                        finalWriter = new ZipArchiveWriter(zipResolver);
                    }
                    else
                    {
                        finalWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);
                    }
                    writer = GetWriter(finalWriter);
                }
                else
                {
                    writer = new XmlTextWriter(new StringWriter());
                }
                source = this.Source;
                // Apply the transformation
                xslt.Transform(source, parameters, writer, zipResolver);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
                if (source != null)
                    source.Close();
                if (zipResolver != null)
                    zipResolver.Dispose();
            }
        }

        private void MessageCallBack(object sender, XsltMessageEncounteredEventArgs e)
        {
            if (e.Message.StartsWith("progress:"))
            {
                if (progressMessageIntercepted != null)
                {
                    progressMessageIntercepted(this, null);
                }
            }
            else if (e.Message.StartsWith("translation.odf2oox."))
            {
                if (feedbackMessageIntercepted != null)
                {
                    feedbackMessageIntercepted(this, new OdfEventArgs(e.Message));
                }
            }
            else if (e.Message.StartsWith("translation.oox2odf."))
            {
                if (feedbackMessageIntercepted != null)
                {
                    feedbackMessageIntercepted(this, new OdfEventArgs(e.Message));
                }
            }
        }

        private void CheckFile(string fileName)
        {
            if (this.isDirectTransform)
            {
                CheckOdfFile(fileName);
            }
            else
            {
                CheckOoxFile(fileName);
            }
        }

        /// <summary>
        /// Test if the input file is an ODF document.
        /// Throw NotAndOdfDocumentException and/or EncryptedDocumentException
        /// </summary>
        /// <param name="fileName">input file name</param>
        protected abstract void CheckOdfFile(string fileName);

        /// <summary>
        /// Test if the input file is an OOX document
        /// Throw NotAndOoxDocumentException 
        /// </summary>
        /// <param name="fileName"></param>
        protected abstract void CheckOoxFile(string fileName);


        private XmlWriter GetWriter(XmlWriter writer)
        {
            string [] postProcessors = this.DirectPostProcessorsChain;
            if (!this.isDirectTransform)
            {
                postProcessors = this.ReversePostProcessorsChain;
            }
            return InstanciatePostProcessors(postProcessors, writer);
        }

        private XmlWriter InstanciatePostProcessors(string [] procNames, XmlWriter lastProcessor)
        {
            XmlWriter currentProc = lastProcessor;
            for (int i = procNames.Length - 1; i >= 0; --i)
            {
                if (!this.skipedPostProcessors.Contains(procNames[i]))
                {
                    Type type = Type.GetType("CleverAge.OdfConverter.OdfConverterLib." + procNames[i]);
                    object[] parameters = { currentProc };
                    XmlWriter newProc = (XmlWriter)Activator.CreateInstance(type, parameters);
                    currentProc = newProc;
                }
            }
            return currentProc;
        }
    }
}
