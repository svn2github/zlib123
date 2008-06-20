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


        public Converter()
            : base(Assembly.GetExecutingAssembly())
        { }

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
            catch (Exception)
            {
                return null;
            }
            return stylesheet;
        }
      
        protected override string [] DirectPostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string []  {
                   "OdfConverter.Wordprocessing.OoxChangeTrackingPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.OdfConverterLib.OoxSpacesPostProcessor",
        	       "OdfConverter.Wordprocessing.OoxSectionsPostProcessor,"+fullname, 
        	       "OdfConverter.Wordprocessing.OoxAutomaticStylesPostProcessor,"+fullname,
        	       "OdfConverter.Wordprocessing.OoxParagraphsPostProcessor,"+fullname,
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
                    "OdfConverter.Wordprocessing.OdfParagraphPostProcessor,"+fullname,
			        "OdfConverter.Wordprocessing.OdfCheckIfIndexPostProcessor,"+fullname,
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
