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
using CleverAge.OdfConverter.OdfZipUtils;
using System.Collections.Generic;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    /// Core conversion methods 
    /// </summary>
    public abstract class AbstractConverter
    {
        protected const string ODFToOOX_XSL = "odf2oox.xsl";
        protected const string OOXToODF_XSL = "oox2odf.xsl";
        protected const string SOURCE_XML = "source.xml";
        protected const string ODFToOOX_COMPUTE_SIZE_XSL = "odf2oox-compute-size.xsl";
        protected const string OOXToODF_COMPUTE_SIZE_XSL = "oox2odf-compute-size.xsl";
      
        protected bool isDirectTransform = true;
        protected List<string> skippedPostProcessors = null;
        protected string externalResource = null;
        protected bool packaging = true;
        protected Assembly resourcesAssembly;
        protected Dictionary<string, XslCompiledTransform> compiledProcessors;
        //Added by Sonata-15/11/2007   
        //static varibale is used for getting temporary input file name
        public static string inputTempFileName;
        
        /// <summary>
        /// Derived classed may return a precompiled stylesheet type.
        /// </summary>
        protected virtual Type LoadPrecompiledXslt()
        {
            return null;
        }
                
        protected AbstractConverter(Assembly resourcesAssembly)
        {
            this.resourcesAssembly = resourcesAssembly;
            this.skippedPostProcessors = new List<string>();
            this.compiledProcessors = new Dictionary<string, XslCompiledTransform>();
        }

        public bool DirectTransform
        {
            set { this.isDirectTransform = value; }
            get { return this.isDirectTransform; }
        }

        public List<string> SkippedPostProcessors
        {
            get { return this.skippedPostProcessors; }
            set { this.skippedPostProcessors = value; }
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
        /// Pull the chain of post processors for the direct conversion
        /// </summary>
        protected virtual string [] DirectPostProcessorsChain
        {
            get { return null; }
        }

        /// <summary>
        /// Pull the chain of post processors for the reverse conversion
        /// </summary>
        protected virtual string [] ReversePostProcessorsChain
        {
            get { return null; }
        }

        /// <summary>
        /// Pull an XmlUrlResolver for embedded resources
        /// </summary>
        protected XmlUrlResolver ResourceResolver
        {
            get
            {
                if (this.ExternalResources == null)
                {
                    return new EmbeddedResourceResolver(this.resourcesAssembly,
                      this.GetType().Namespace, this.DirectTransform);
                }
                else
                {
                    return new SharedXmlUrlResolver(this.DirectTransform);
                }
            }
        }

        /// <summary>
        /// Pull the input xml document to the xsl transformation
        /// </summary>
        protected virtual XmlReader Source(string inputFile)
        {
            XmlReaderSettings xrs = new XmlReaderSettings();
            // do not look for DTD
            xrs.ProhibitDtd = false;
            if (this.ExternalResources == null)
            {
                xrs.XmlResolver = this.ResourceResolver;
                return XmlReader.Create(SOURCE_XML, xrs);
            }
            else
            {
                return XmlReader.Create(this.ExternalResources + "/" + SOURCE_XML, xrs);
            }
        }

      
        protected virtual XslCompiledTransform Load(bool computeSize)
        {
            string xslLocation = this.DirectTransform ? ODFToOOX_XSL : OOXToODF_XSL;
            XPathDocument xslDoc = null;
            XmlUrlResolver resolver = this.ResourceResolver;

            if (this.ExternalResources == null)
            {
                if (computeSize)
                {
                    xslLocation = this.DirectTransform ? ODFToOOX_COMPUTE_SIZE_XSL : OOXToODF_COMPUTE_SIZE_XSL;
                }
                EmbeddedResourceResolver emr = (EmbeddedResourceResolver)resolver;
                emr.IsDirectTransform = this.DirectTransform;
                xslDoc = new XPathDocument(emr.GetInnerStream(xslLocation));
            }
            else
            {
                xslDoc = new XPathDocument(this.ExternalResources + "/" + xslLocation);
            }

            if (!this.compiledProcessors.ContainsKey(xslLocation))
            {
                // create an XSL transformer

                // Activation of XSL Debugging only in "DEBUG" compilation mode
#if DEBUG
                XslCompiledTransform xslt = new XslCompiledTransform(true);
#else
                XslCompiledTransform xslt = new XslCompiledTransform();
#endif

                //JP - 03/07/2007
                // compile the stylesheet. 
                // Input stylesheet, xslt settings and uri resolver are retrieve from the implementation class.
                try
                {
#if (!DEBUG)
                    Type t = typeof(XslCompiledTransform);
                    MethodInfo mi = t.GetMethod("Load", new Type[] { typeof(Type) });
                    Type compiledStylesheet = this.LoadPrecompiledXslt();

                    // check if optimization of precompiled XSLT works on current 
                    // .NET Framework installation and with current conversion direction
                    // (this feature requires .NET Framework 2.0 SP1)
                    if (mi != null && compiledStylesheet != null)
                    {
                        // dynamically invoke xslt.Load(compiledStylesheet);
                        mi.Invoke(xslt, new object[] { compiledStylesheet });
                    }
                    else
                    {
#endif
                        xslt.Load(xslDoc, this.XsltProcSettings, this.ResourceResolver);
#if (!DEBUG)
                    }
#endif
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                this.compiledProcessors.Add(xslLocation, xslt);
            }
            return (XslCompiledTransform)this.compiledProcessors[xslLocation];
        }

        /// <summary>
        /// Pull the xslt settings
        /// </summary>
        protected XsltSettings XsltProcSettings
        {
            get
            {
                // Enable xslt 'document()' function
                return new XsltSettings(true, false);
            }
        }

        /// <summary>
        /// Test if the input file is an ODF document.
        /// Throw NotAndOdfDocumentException and/or EncryptedDocumentException
        /// </summary>
        /// <param name="fileName">input file name</param>
        protected virtual void CheckOdfFile(string fileName) { }

        /// <summary>
        /// Test if the input file is an OOX document
        /// Throw NotAndOoxDocumentException 
        /// </summary>
        /// <param name="fileName"></param>
        protected virtual void CheckOoxFile(string fileName) { }




        public void ComputeSize(string inputFile)
        {
            Transform(inputFile, null);
        }

        public void Transform(string inputFile, string outputFile)
        {
            Transform(inputFile, outputFile, null);
        }

        /// <summary>
        /// bug #1644285 Zlib crashes on non-ascii file names.
        /// TODO : temp folders
        /// </summary>
        public void Transform(string inputFile, string outputFile, ConversionOptions options)
        {
            string tempInputFile = Path.GetTempFileName();
            string tempOutputFile = outputFile == null ? null : Path.GetTempFileName();
            try
            {
                File.Copy(inputFile, tempInputFile, true);
                File.SetAttributes(tempInputFile, FileAttributes.Normal);
                //Added by sonata -15/11/2007              
                inputTempFileName = tempInputFile;
                //End
                _Transform(tempInputFile, tempOutputFile, options);

                if (outputFile != null)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    // make sure that the output folder exists
                    FileInfo fi = new FileInfo(outputFile);
                    Directory.CreateDirectory(fi.DirectoryName);

                    File.Move(tempOutputFile, outputFile);
                }
            }
            finally
            {
                if (File.Exists(tempInputFile))
                {
                    try
                    {
                        File.Delete(tempInputFile);
                    }
                    catch (IOException)
                    {
                        Debug.Write("could not delete temporary input file");
                    }
                }
            }
        }

        
        protected virtual void _Transform(string inputFile, string outputFile, ConversionOptions options)
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
                XslCompiledTransform xslt =  this.Load(outputFile == null);
                zipResolver = new ZipResolver(inputFile);
                XsltArgumentList parameters = new XsltArgumentList();
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                
                if (outputFile != null)
                {
                    parameters.AddParam("outputFile", "", outputFile);
                    if (options != null)
                    {
                        parameters.AddParam("documentType", "", options.DocumentType.ToString());
                        parameters.AddParam("generator", "", options.Generator);
                    }
                    else
                    {
                        parameters.AddParam("generator", "", "OpenXML/ODF Translator v2.5");
                    }
                    XmlWriter finalWriter;
                    if (this.packaging)
                    {
                        finalWriter = new ZipArchiveWriter(zipResolver, outputFile);
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
                source = this.Source(inputFile);
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

        protected void MessageCallBack(object sender, XsltMessageEncounteredEventArgs e)
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

        protected void CheckFile(string fileName)
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

  

        protected XmlWriter GetWriter(XmlWriter writer)
        {
            string [] postProcessors = this.DirectPostProcessorsChain;
            if (!this.isDirectTransform)
            {
                postProcessors = this.ReversePostProcessorsChain;
            }
            return InstanciatePostProcessors(postProcessors, writer);
        }


        protected XmlWriter InstanciatePostProcessors(string [] procNames, XmlWriter lastProcessor)
        {
            XmlWriter currentProc = lastProcessor;
            if (procNames != null)
            {
                for (int i = procNames.Length - 1; i >= 0; --i)
                {
                    if (!Contains(procNames[i], this.skippedPostProcessors))
                    {
                        Type type = Type.GetType(procNames[i]);
                        object [] parameters = { currentProc };
                        XmlWriter newProc = (XmlWriter) Activator.CreateInstance(type, parameters);
                        currentProc = newProc;
                    }
                }
            }
            return currentProc;
        }

        protected bool Contains(string processorFullName, List<string> names)
        {
            foreach (string name in names)
            {
                if (processorFullName.Contains(name))
                {
                    return true;
                }
            }
            return false;
        }

        public delegate void MessageListener(object sender, EventArgs e);
        protected event MessageListener progressMessageIntercepted;
        protected event MessageListener feedbackMessageIntercepted;

        public void AddProgressMessageListener(MessageListener listener)
        {
            progressMessageIntercepted += listener;
        }

        public void AddFeedbackMessageListener(MessageListener listener)
        {
            feedbackMessageIntercepted += listener;
        }

        public void RemoveMessageListeners()
        {
            progressMessageIntercepted = null;
            feedbackMessageIntercepted = null;
        }

    }
}
