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
        private const string ODFToOOX_LOCATION = "odf2oox";
        private const string OOXToODF_LOCATION = "oox2odf";
        private const string ODFToOOX_XSL = "odf2oox.xsl";
        private const string OOXToODF_XSL = "oox2odf.xsl";
        private const string SOURCE_XML = "source.xml";
        private const string ODF_TEXT_MIME = "application/vnd.oasis.opendocument.text";
        
        private EmbeddedResourceResolver _embeddedResolver;

    
        protected override string[] DirectPostProcessorsChain
        {
            get
            {
                return new string []  {
                    "OoxChangeTrackingPostProcessor",
                    "OoxSpacesPostProcessor",
        	        "OoxSectionsPostProcessor", 
        	        "OoxAutomaticStylesPostProcessor",
        	        "OoxParagraphsPostProcessor",
        	        "OoxCharactersPostProcessor" 
                };
            }
        }

        protected override string[] ReversePostProcessorsChain
        {
            get
            {
                return new string []  {
                    "OdfParagraphPostProcessor",
			        "OdfCheckIfIndexPostProcessor",
        	        "OdfCharactersPostProcessor"
                };
            }
        }


        protected override XmlUrlResolver ResourceResolver
        {
            get
            {
                if (this.ExternalResources == null)
                {
                    if (this._embeddedResolver == null)
                    {
                        string resourcesLocation = this.DirectTransform ? ODFToOOX_LOCATION : OOXToODF_LOCATION;
                        this._embeddedResolver =  new EmbeddedResourceResolver(Assembly.GetExecutingAssembly(),
                            this.GetType().Namespace + ".resources." + resourcesLocation, 
                            this.DirectTransform);
                    }
                    return this._embeddedResolver;
                }
                else
                {
                    return new SharedXmlUrlResolver(this.DirectTransform);
                }
            }
        }

        protected override XmlReader Source
        {
            get
            {
                XmlReaderSettings xrs = new XmlReaderSettings();
                // do not look for DTD
                xrs.ProhibitDtd = true;
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
        }

        /// <summary>
        /// Get the input xsl document to the xsl transformation
        /// </summary>
        protected override XPathDocument XslDoc
        {
            get {
                string xslLocation = this.DirectTransform ? ODFToOOX_XSL : OOXToODF_XSL;
                if (this.ExternalResources == null)
                {
                    return new XPathDocument(((EmbeddedResourceResolver) 
                        this.ResourceResolver).GetInnerStream(xslLocation));
                }
                else
                {
                    return new XPathDocument(this.ExternalResources + "/" + xslLocation);
                }
            }
        }

        protected override XsltSettings XsltProcSettings
        {
            get 
            {
                // Enable xslt 'document()' function
                return new XsltSettings(true, false); 
            }
        }

        protected override void CheckOoxFile(string fileName) 
        {
            // TODO: implement
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
                                                + ODF_TEXT_MIME + "']", nsmgr);
            if (node == null)
            {
                throw new NotAnOdfDocumentException("Could not convert " + fileName
                                                    + ". Invalid OASIS OpenDocument file");
            }
        }
    
    }
}
