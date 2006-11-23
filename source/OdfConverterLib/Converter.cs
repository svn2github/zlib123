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
        private const string ODF_MIME_TYPE = "application/vnd.oasis.opendocument.text";

        private string[] OOX_POST_PROCESSORS = 
        {
            "OoxChangeTrackingPostProcessor",
        	"OoxSectionsPostProcessor", 
        	"OoxAutomaticStylesPostProcessor",
        	"OoxParagraphsPostProcessor",
        	"OoxCharactersPostProcessor" 
        };

        private string[] ODF_POST_PROCESSORS = {
			"OdfParagraphPostProcessor"
 		};

        private bool isDirectTransform = true;
        private ArrayList skipedPostProcessors = null;
        private string externalResource = null;
        private bool packaging = true;

        public Converter()
        {
            this.skipedPostProcessors = new ArrayList();
        }

        public bool DirectTransform
        {
            set { this.isDirectTransform = value; }
        }

        public ArrayList SkipedPostProcessors
        {
            set { this.skipedPostProcessors = value; }
        }

        public string ExternalResources
        {
            set { this.externalResource = value; }
        }

        public bool Packaging
        {
            set { this.packaging = value; }
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

        public void Transform(string inputFile, string outputFile)
        {
            // this throws an exception in the the following cases:
            // - input file is not a valid file
            // - input file is an encrypted file
            CheckFile(inputFile);

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
                string resourcesLocation = ODFToOOX_LOCATION;
                string xslLocation = ODFToOOX_XSL;
                if (!this.isDirectTransform)
                {
                    resourcesLocation = OOXToODF_LOCATION;
                    xslLocation = OOXToODF_XSL;
                }

                if (this.externalResource == null)
                {

                    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + RESOURCE_LOCATION + "." + resourcesLocation);
                    xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(xslLocation));
                    xrs.XmlResolver = resourceResolver;
                    source = XmlReader.Create(SOURCE_XML, xrs);

                }
                else
                {
                    resourceResolver = new XmlUrlResolver();
                    xslDoc = new XPathDocument(this.externalResource + "/" + xslLocation);
                    source = XmlReader.Create(this.externalResource + "/" + SOURCE_XML, xrs);
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

        private void CheckOdfFile(string fileName)
        {
            // Test for encryption
            XmlDocument doc;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.XmlResolver = new ZipResolver(fileName);
                settings.ProhibitDtd = false;
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
            
            // Check the document mime-type.
      		XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
      		nsmgr.AddNamespace("manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
           
            XmlNode node = doc.SelectSingleNode("/manifest:manifest/manifest:file-entry[@manifest:media-type='"
                                                + ODF_MIME_TYPE + "']", nsmgr);
           	if (node == null) 
            {
            	throw new NotAnOdfDocumentException("Could not convert "+ fileName 
           		                                    + ". Invalid OASIS OpenDocument file");
            }        
        }

        private void CheckOoxFile(string fileName)
        {
            // TODO: implement
        }

        private XmlWriter GetWriter(XmlWriter writer)
        {
            string [] postProcessors = OOX_POST_PROCESSORS;
            if (!this.isDirectTransform)
            {
                postProcessors = ODF_POST_PROCESSORS;
            }
            return InstanciatePostProcessors(postProcessors, writer);
        }

        private XmlWriter InstanciatePostProcessors(string[] procNames, XmlWriter lastProcessor)
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
