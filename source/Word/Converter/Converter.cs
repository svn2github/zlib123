/* 
 * Copyright (c) 2006-2008 Clever Age
 * Copyright (c) 2006-2009 DIaLOGIKa
 *
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
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using CleverAge.OdfConverter.OdfConverterLib;
using CleverAge.OdfConverter.OdfZipUtils;
using System.IO;


namespace OdfConverter.Wordprocessing
{
    public class Converter : AbstractConverter
    {

        private const string ODF_TEXT_MIME = "application/vnd.oasis.opendocument.text";
        private const string ODF_TEXT_TEMPLATE_MIME = "application/vnd.oasis.opendocument.text-template";

        public Converter()
            : base(Assembly.GetExecutingAssembly())
        {
        }

        protected override Type LoadPrecompiledXslt()
        {
            Type stylesheet = null;
            try
            {
                if (this.DirectTransform)
                {
                    stylesheet = Assembly.Load("WordprocessingConverter2Oox")
                                            .GetType("WordprocessingConverter2Oox");
                }
                else
                {
                    stylesheet = Assembly.Load("WordprocessingConverter2Odf")
                                            .GetType("WordprocessingConverter2Odf");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                return null;
            }
            return stylesheet;
        }

        protected override string[] DirectPostProcessorsChain
        {
            get
            {
                //ODF -> DOCX

                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string[]  {
                   "OdfConverter.Wordprocessing.OoxChangeTrackingPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.OdfConverterLib.OoxSpacesPostProcessor",
                   "OdfConverter.Wordprocessing.OoxSectionsPostProcessor,"+fullname, 
                   "OdfConverter.Wordprocessing.OoxAutomaticStylesPostProcessor,"+fullname,
                   "OdfConverter.Wordprocessing.OoxParagraphsPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.OdfConverterLib.OoxCharactersPostProcessor", 
                   
                };
                // "OdfConverter.Wordprocessing.OoxReplacementPostProcessor,"+fullname
            }
        }

        protected override string[] ReversePostProcessorsChain
        {
            get
            {
                //DOCX -> ODF

                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string[]  {
                    "OdfConverter.Wordprocessing.OdfParagraphPostProcessor,"+fullname,
                    //"OdfConverter.Wordprocessing.OdfCheckIfIndexPostProcessor,"+fullname,
                    "CleverAge.OdfConverter.OdfConverterLib.OdfCharactersPostProcessor",
                    "OdfConverter.Wordprocessing.OdfIndexSourceStylesPostProcessor,"+fullname
                };
            }
        }


        protected override void CheckOdfFile(string fileName)
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
            catch (XmlException ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString()); 
                throw new NotAnOdfDocumentException(ex.Message, ex);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                throw;
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
                                                + ODF_TEXT_MIME + "']", nsmgr);

            XmlNodeList mediaTypeNodes = doc.SelectNodes(
                "/manifest:manifest/manifest:file-entry[@manifest:media-type='" + ODF_TEXT_MIME + "' or @manifest:media-type='" + ODF_TEXT_TEMPLATE_MIME + "']", nsmgr);
            if (mediaTypeNodes.Count == 0)
            {
                throw new NotAnOdfDocumentException("Could not convert " + fileName + ". Invalid OASIS OpenDocument file");
            }
        }

        /// <summary>
        /// Pull the input xml document to the xsl transformation
        /// </summary>
        protected override XmlReader Source(string inputFile)
        {
            if (this.DirectTransform)
            {
                // ODT -> DOCX
                return base.Source(inputFile);
            }
            else
            {
                // use performance improved method for
                // DOCX -> ODT conversion
                XmlReaderSettings xrs = new XmlReaderSettings();
                // do not look for DTD
                xrs.ProhibitDtd = false;

                DocxDocument doc = new DocxDocument(inputFile);

                // uncomment for testing
                for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
                {
                    if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/DUMP")
                    {
                        Stream package = doc.OpenXML;
                        FileInfo fi = new FileInfo(Environment.GetCommandLineArgs()[i + 1]);
                        Stream s = fi.OpenWrite();
                        byte[] buffer = new byte[package.Length];
                        package.Read(buffer, 0, (int)package.Length);
                        s.Write(buffer, 0, (int)package.Length);
                        s.Close();
                    }
                }

                return XmlReader.Create(doc.OpenXML, xrs);
            }
        }
    }
}
