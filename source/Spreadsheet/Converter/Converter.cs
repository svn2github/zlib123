using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;
using CleverAge.OdfConverter.OdfConverterLib;


namespace CleverAge.OdfConverter.Spreadsheet
{
    public class Converter : AbstractConverter
    {

        private const string ODF_TEXT_MIME = "application/vnd.oasis.opendocument.spreadsheet";
        private const string OOX_TEXT_CONTENTTYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";


        public Converter()
            : base(Assembly.GetExecutingAssembly())
        { }
      
        protected override string [] DirectPostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string []  {
                   "CleverAge.OdfConverter.Spreadsheet.OoxCommentsPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.Spreadsheet.OoxDrawingsPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.Spreadsheet.OoxHeaderFooterPostProcessor,"+fullname,
				   "CleverAge.OdfConverter.Spreadsheet.OOXStyleCellPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.Spreadsheet.OoxMaximumCellTextPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.Spreadsheet.OoxPhysicalPathPostProcessor,"+fullname,
                   "CleverAge.OdfConverter.Spreadsheet.OoxPivotCachePostProcessor,"+fullname,
                };
            }
        }
        protected override string [] ReversePostProcessorsChain
        {
            get
            {
                string fullname = Assembly.GetExecutingAssembly().FullName;
                return new string []  {
                    "CleverAge.OdfConverter.Spreadsheet.OdfSharedStringsPostProcessor,"+fullname,
                    "CleverAge.OdfConverter.Spreadsheet.OOXGroupsPostProcessor,"+fullname
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
                throw new NotAnOdfDocumentException(e.Message);
            }
            catch (Exception e)
            {
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

       // Code to fix the bug# 1698280 

        protected override void CheckOoxFile(string fileName)
        {
            
            XmlDocument doc;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.XmlResolver = new ZipResolver(fileName);
                settings.ProhibitDtd = false;
                doc = new XmlDocument();
                XmlReader reader = XmlReader.Create("[Content_Types].xml", settings);
                doc.Load(reader);
            }
            catch (XmlException e)
            {
                throw new NotAnOoxDocumentException(e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }
            XmlNodeList nodelist = doc.SelectNodes("//node()[@ContentType='" + OOX_TEXT_CONTENTTYPE + "']");
            if (nodelist.Count== 0)
            {
                throw new NotAnOoxDocumentException("not an valid oox file");
            }
                  
        }

        // End

    }
}
