using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using CleverAge.OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.Word
{
    public class Converter : AbstractConverter
    {

        private const string ODF_TEXT_MIME = "application/vnd.oasis.opendocument.text";


        public Converter()
            : base(Assembly.GetExecutingAssembly())
        { }
      
        protected override string [] DirectPostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string []  {
                   "CleverAge.OdfConverter.Word.OoxChangeTrackingPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.OdfConverterLib.OoxSpacesPostProcessor",
        	       "CleverAge.OdfConverter.Word.OoxSectionsPostProcessor,"+fullname, 
        	       "CleverAge.OdfConverter.Word.OoxAutomaticStylesPostProcessor,"+fullname,
        	       "CleverAge.OdfConverter.Word.OoxParagraphsPostProcessor,"+fullname,
        	       "CleverAge.OdfConverter.OdfConverterLib.OoxCharactersPostProcessor"
                };
            }
        }

        protected override string [] ReversePostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string []  {
                    "CleverAge.OdfConverter.Word.OdfParagraphPostProcessor,"+fullname,
			        "CleverAge.OdfConverter.Word.OdfCheckIfIndexPostProcessor,"+fullname,
        	        "CleverAge.OdfConverter.OdfConverterLib.OdfCharactersPostProcessor",
                    "CleverAge.OdfConverter.Word.OdfIndexSourceStylesPostProcessor,"+fullname
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
            catch (XmlException e)
            {
                Debug.WriteLine(e.Message);
                throw new NotAnOdfDocumentException(e.Message);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                throw e;
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
    
    }
}
