using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Xml;
using System.Diagnostics;
using CleverAge.OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    /// An abstract base class for converting files according to Open Packaging Conventions standard
    /// to a single XML stream.
    /// This XML stream has the following form:
    ///     &lt;?xml version="1.0" encoding="utf-8" standalone="yes"?&gt;
    ///     &lt;oox:package xmlns:oox="urn:oox"&gt;
    ///         &lt;oox:part oox:name="_rels/.rels"&gt;
    ///             [content of package part _rels/.rels]
    ///         &lt;/oox:part&gt;
    ///         &lt;oox:part oox:name="docProps/app.xml" oox:type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" oox:rId="rId3"&gt;
    ///             [content of package part docProps/app.xml]
    ///         &lt;/oox:part&gt;
    ///         &lt;oox:part oox:name="docProps/core.xml" oox:type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" oox:rId="rId2"&gt;
    ///             [content of package part docProps/core.xml]
    ///         &lt;/oox:part&gt;
    ///         &lt;oox:part oox:name="word/_rels/document.xml.rels"&gt;
    ///             [content of package part word/_rels/document.xml.rels]
    ///         &lt;/oox:part&gt;
    ///         &lt;oox:part oox:name="word/document.xml" oox:type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" oox:rId="rId1"&gt;  
    ///             [content of package part word/document.rels]
    ///         &lt;/oox:part&gt;
    ///     &lt;/oox:package>
    /// </summary>
    public abstract class OoxDocument : IDisposable
    {
        protected struct RelationShip
        {
            public string Id;
            public string Target;
            public string Type;
        }

        protected string _fileName;
        protected Stream _stream;
        private bool _disposed = false;

        // use a similar schema than defined by "http://schemas.microsoft.com/office/2006/xmlPackage"
        protected const string PACKAGE_NS = "urn:oox";
        protected const string NS_PREFIX = "oox";
        protected const string RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

        public OoxDocument(string fileName)
        {
            this._fileName = fileName;
        }

        ~OoxDocument()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        public void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_stream != null)
                        _stream.Close();
                }
                _disposed = true;
            }
        }

        /// <summary>
        /// A stream containing all XML parts of the document whose namespaces are defined in property <code>Namespaces</code>.
        /// </summary>
        public Stream OpenXML
        {
            get
            {
                try
                {
                    ZipReader archive = ZipFactory.OpenArchive(_fileName);
                    
                    MemoryStream ms = new MemoryStream();
                    XmlTextWriter xtw = new XmlTextWriter(new StreamWriter(ms, Encoding.UTF8));
                    
                    xtw.WriteStartDocument(true);
                    //xtw.WriteProcessingInstruction("mso-application", "progid='Word.Document'");
                    xtw.WriteStartElement(NS_PREFIX, "package", PACKAGE_NS);

                    // copy the whole document from its root
                    CopyLevel(archive, "_rels/.rels", xtw, this.Namespaces);
                    
                    xtw.WriteEndDocument();
                    xtw.Flush();

                    // reset stream
                    ms.Seek(0L, SeekOrigin.Begin);
                    
                    return ms;
                }
                catch (Exception ex)
                {
                    throw new NotAnOoxDocumentException("Could not create XML stream from input document.", ex);
                }
            }
        }

        protected abstract void CopyLevel(ZipReader archive, string relFile, XmlTextWriter xtw, List<string> namespaces);

        /// <summary>
        /// A list of namespaces to be copied from the archive into the intermediate XML format
        /// </summary>
        protected abstract List<string> Namespaces
        {
            get;
        }
    }
}
