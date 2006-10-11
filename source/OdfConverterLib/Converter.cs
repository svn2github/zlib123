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
    public class Converter
    {
        private const string RESOURCE_LOCATION = "resources";
        private const string ODFToOOX_LOCATION = "odf2oox";
        private const string ODFToOOX_XSL = "odf2oox.xsl";
        private const string ODFToOOX_COMPUTE_SIZE_XSL = "odf2oox-compute-size.xsl";
        private const string OOXToODF_LOCATION = "oox2odf";
        private const string OOXToODF_XSL = "oox2odf.xsl";
        private const string OOXToODF_COMPUTE_SIZE_XSL = "oox2odf-compute-size.xsl";
        private const string SOURCE_XML = "source.xml";

        private string[] OOX_POST_PROCESSORS = { "OoxAutomaticStylesPostProcessor", "OoxCharactersPostProcessor" };
        private string[] ODF_POST_PROCESSORS = { };

        public Converter()
        {
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

        public void OdfToOox(string inputFile, string outputFile)
        {
            DoOdfToOoxTransform(inputFile, outputFile, null, null);
        }

        public void OdfToOoxSkipPostProcess(string inputFile, string outputFile, string skipPostProcessor)
        {
            DoOdfToOoxTransform(inputFile, outputFile, null, skipPostProcessor);
        }

        public void OdfToOoxWithExternalResources(string inputFile, string outputFile, string resourceDir)
        {
            DoOdfToOoxTransform(inputFile, outputFile, resourceDir, null);
        }

        public void OdfToOoxSkipPostProcessWithExternalResources(string inputFile, string outputFile, string resourceDir, string skipPostProcessor)
        {
            DoOdfToOoxTransform(inputFile, outputFile, resourceDir, skipPostProcessor);
        }

        public void OdfToOoxComputeSize(string inputFile)
        {
            DoOdfToOoxTransform(inputFile, null, null, null);
        }

        public void OoxToOdf(string inputFile, string outputFile)
        {
            DoOoxToOdfTransform(inputFile, outputFile, null, null);
        }

        public void OoxToOdfSkipPostProcess(string inputFile, string outputFile, string skipPostProcessor)
        {
            DoOoxToOdfTransform(inputFile, outputFile, null, skipPostProcessor);
        }

        public void OoxToOdfWithExternalResources(string inputFile, string outputFile, string resourceDir)
        {
            DoOoxToOdfTransform(inputFile, outputFile, resourceDir, null);
        }

        public void OoxToOdfSkipPostProcessWithExternalResources(string inputFile, string outputFile, string resourceDir, string skipPostProcessor)
        {
            DoOoxToOdfTransform(inputFile, outputFile, resourceDir, skipPostProcessor);
        }

        public void OoxToOdfComputeSize(string inputFile)
        {
            DoOoxToOdfTransform(inputFile, null, null, null);
        }

        private void DoOdfToOoxTransform(string inputFile, string outputFile, string resourceDir, string skipPostProcessor)
        {
            // this throws an exception in the the following cases:
            // - input file is not an ODF file
            // - input file is an encrypted ODF file
            CheckOdf(inputFile);

            XmlUrlResolver resourceResolver;
            XPathDocument xslDoc;
            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader source = null;
            XmlWriter writer = null;
            ZipResolver zipResolver = null;

            try
            {

                // do not look for DTD
                xrs.ProhibitDtd = true;

                if (resourceDir == null)
                {
                    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + RESOURCE_LOCATION + "." + ODFToOOX_LOCATION);
                    xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(ODFToOOX_XSL));
                    xrs.XmlResolver = resourceResolver;
                    source = XmlReader.Create(SOURCE_XML, xrs);

                }
                else
                {
                    resourceResolver = new XmlUrlResolver();
                    xslDoc = new XPathDocument(resourceDir + "/" + ODFToOOX_XSL);
                    source = XmlReader.Create(resourceDir + "/" + SOURCE_XML, xrs);
                }

                // create a xsl transformer
                XslCompiledTransform xslt = new XslCompiledTransform();
                // Enable xslt 'document()' function
                XsltSettings settings = new XsltSettings(true, false);
                // compile the stylesheet
                xslt.Load(xslDoc, settings, resourceResolver);


                zipResolver = new ZipResolver(inputFile);
                XsltArgumentList parameters = new XsltArgumentList();

                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                if (outputFile != null)
                {
                    parameters.AddParam("outputFile", "", outputFile);
                    writer = GetOoxWriter(zipResolver, skipPostProcessor);
                }
                else
                {
                    writer = new XmlTextWriter(new StringWriter());
                }

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

        private void DoOoxToOdfTransform(string inputFile, string outputFile, string resourceDir, string skipPostProcessor)
        {

            // this throws an exception in the the following cases:
            // - input file is not an OOX file
            // - input file is an encrypted OOX file
            CheckOox(inputFile);

            XmlUrlResolver resourceResolver;
            XPathDocument xslDoc;
            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader source = null;
            XmlWriter writer = null;
            ZipResolver zipResolver = null;

            try
            {

                // do not look for DTD
                xrs.ProhibitDtd = true;

                if (resourceDir == null)
                {
                    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + RESOURCE_LOCATION + "." + OOXToODF_LOCATION);
                    xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(OOXToODF_XSL));
                    xrs.XmlResolver = resourceResolver;
                    source = XmlReader.Create(SOURCE_XML, xrs);

                }
                else
                {
                    resourceResolver = new XmlUrlResolver();
                    xslDoc = new XPathDocument(resourceDir + "/" + OOXToODF_XSL);
                    source = XmlReader.Create(resourceDir + "/" + SOURCE_XML, xrs);
                }

                // create a xsl transformer
                XslCompiledTransform xslt = new XslCompiledTransform();
                // Enable xslt 'document()' function
                XsltSettings settings = new XsltSettings(true, false);
                // compile the stylesheet
                xslt.Load(xslDoc, settings, resourceResolver);


                zipResolver = new ZipResolver(inputFile);
                XsltArgumentList parameters = new XsltArgumentList();

                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                if (outputFile != null)
                {
                    parameters.AddParam("outputFile", "", outputFile);
                    writer = GetOdfWriter(zipResolver, skipPostProcessor);
                }
                else
                {
                    writer = new XmlTextWriter(new StringWriter());
                }

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
            else if (e.Message.StartsWith("feedback:"))
            {
                if (feedbackMessageIntercepted != null)
                {
                    feedbackMessageIntercepted(this, new OdfEventArgs(e.Message.Substring("feedback:".Length)));
                }
            }
        }

        private void CheckOdf(string fileName)
        {

            // Test for encryption
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.XmlResolver = new ZipResolver(fileName);
            settings.ProhibitDtd = false;
            XmlDocument doc;
            try
            {
                doc = new XmlDocument();
                XmlReader reader = XmlReader.Create("META-INF/manifest.xml", settings);
                doc.Load(reader);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                throw new NotAnOdfDocumentException(e.Message);
            }
            XmlNodeList nodes = doc.GetElementsByTagName("encryption-data", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
            if (nodes.Count > 0)
            {
                throw new EncryptedDocumentException(fileName + " is an encrypted document");
            }
        }

        private void CheckOox(string fileName)
        {
            // TODO: implement
        }

        private XmlWriter GetOoxWriter(XmlResolver resolver, string skipPostProcessor)
        {
            return InstanciatePostProcessors(OOX_POST_PROCESSORS, new ZipArchiveWriter(resolver), skipPostProcessor);
        }

        private XmlWriter GetOdfWriter(XmlResolver resolver, string skipPostProcessor)
        {
            return InstanciatePostProcessors(ODF_POST_PROCESSORS, new ZipArchiveWriter(resolver), skipPostProcessor);
        }

        private XmlWriter InstanciatePostProcessors(string[] procNames, XmlWriter lastProcessor, string skipPostProcessor)
        {
            XmlWriter currentProc = lastProcessor;
            //XmlWriter currentProc = XmlWriter.Create("output.xml");
            for (int i = procNames.Length - 1; i >= 0; --i)
            {
                if (!procNames[i].Equals(skipPostProcessor))
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
