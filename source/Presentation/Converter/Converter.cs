/* 
 * Copyright (c) 2006, Sonata Software Limited
 * All rights reserved.
 * 
 
 */
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using CleverAge.OdfConverter.OdfConverterLib;

namespace Sonata.OdfConverter.Presentation
{
    public class Converter : AbstractConverter
    {
        private readonly string ODF_TEXT_MIME = "application/vnd.oasis.opendocument.presentation";
       
        public Converter() : base(Assembly.GetExecutingAssembly()) { }

        protected override string[] DirectPostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string[]  {
                   "CleverAge.OdfConverter.OdfConverterLib.OoxSpacesPostProcessor",
        	       "CleverAge.OdfConverter.OdfConverterLib.OoxCharactersPostProcessor"
                };
            }
        }

        protected override string[] ReversePostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string[]  {
                    "CleverAge.OdfConverter.OdfConverterLib.OdfCharactersPostProcessor"
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
            catch (Exception e)
            {
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
                                                + ODF_TEXT_MIME + "']", nsmgr);
            if (node == null)
            {
                throw new NotAnOdfDocumentException("Could not convert " + fileName
                                                    + ". Invalid OASIS OpenDocument file");
            }
        }

        protected override void CheckOoxFile(string fileName)
        {
            //Validator for pptx
            PptxValidator v = new PptxValidator();
            v.validate(fileName); 
        }

    }
}
