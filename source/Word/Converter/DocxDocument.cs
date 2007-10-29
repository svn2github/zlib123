using System;
using System.Collections.Generic;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Xml;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Diagnostics;

namespace CleverAge.OdfConverter.Word
{
    class DocxDocument : OoxDocument
    {
        protected int _paraId = 0;
        protected int _sectPrId = 0;

        public DocxDocument(string fileName) : base(fileName)
        {
        }
        
        protected override void CopyLevel(ZipReader archive, string relFile, XmlTextWriter xtw, List<string> namespaces)
        {
            XmlReader reader;
            List<RelationShip> rels;
            XmlReaderSettings xrs = new XmlReaderSettings();
            xrs.IgnoreWhitespace = false;

            try
            {
                // copy relationship part and extract relationships
                reader = XmlReader.Create(archive.GetEntry(relFile));

                xtw.WriteStartElement(NS_PREFIX, "part", PACKAGE_NS);
                xtw.WriteAttributeString(NS_PREFIX, "name", PACKAGE_NS, relFile);
                xtw.WriteAttributeString(NS_PREFIX, "type", PACKAGE_NS, RELATIONSHIP_NS);
                rels = Copy(reader, xtw, true);
                reader.Close();
                xtw.WriteEndElement();
                xtw.Flush();

                // init id counters
                if (relFile.Equals("_rels/.rels"))
                {
                    _paraId = 0;
                    _sectPrId = 0;
                }

                // copy referenced parts
                foreach (RelationShip rel in rels)
                {
                    if (namespaces.Contains(rel.Type))
                    {
                        try
                        {
                            string basePath = relFile.Substring(0, relFile.LastIndexOf("_rels/"));
                            string path = rel.Target.Substring(0, rel.Target.LastIndexOf('/') + 1);
                            string file = rel.Target.Substring(rel.Target.LastIndexOf('/') + 1);

                            // if there is a relationship part for the current part 
                            // copy relationships and referenced files recursively
                            CopyLevel(archive, basePath + path + "_rels/" + file + ".rels", xtw, namespaces);

                            reader = XmlReader.Create(archive.GetEntry(basePath + rel.Target));

                            xtw.WriteStartElement(NS_PREFIX, "part", PACKAGE_NS);
                            xtw.WriteAttributeString(NS_PREFIX, "name", PACKAGE_NS, basePath + rel.Target);
                            xtw.WriteAttributeString(NS_PREFIX, "type", PACKAGE_NS, rel.Type);
                            xtw.WriteAttributeString(NS_PREFIX, "rId", PACKAGE_NS, rel.Id);

                            Copy(reader, xtw, false);
                            _paraId++;
                            _sectPrId++;

                            reader.Close();

                            xtw.WriteEndElement();
                            xtw.Flush();
                        }
                        catch (ZipEntryNotFoundException)
                        {
                        }
                    }
                }
            }
            catch (ZipEntryNotFoundException)
            {
            }
        }

        protected List<RelationShip> Copy(XmlReader xtr, XmlTextWriter xtw, bool extractRels)
        {
            bool isInRel = false;
            List<RelationShip> rels = new List<RelationShip>();

            RelationShip rel = new RelationShip();

            while (xtr.Read())
            {
                switch (xtr.NodeType)
                {
                    case XmlNodeType.Attribute:
                        break;
                    case XmlNodeType.CDATA:
                        xtw.WriteCData(xtr.Value);
                        break;
                    case XmlNodeType.Comment:
                        xtw.WriteComment(xtr.Value);
                        break;
                    case XmlNodeType.DocumentType:
                        xtw.WriteDocType(xtr.Name, null, null, null);
                        break;
                    case XmlNodeType.Element:
                        if (extractRels && xtr.LocalName == "Relationship" && xtr.NamespaceURI == RELATIONSHIP_NS)
                        {
                            isInRel = true;
                            rel = new RelationShip();
                        }

                        xtw.WriteStartElement(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI);

                        if (xtr.HasAttributes)
                        {
                            while (xtr.MoveToNextAttribute())
                            {
                                if (extractRels && isInRel)
                                {
                                    if (xtr.LocalName == "Type")
                                        rel.Type = xtr.Value;
                                    else if (xtr.LocalName == "Target")
                                        rel.Target = xtr.Value;
                                    else if (xtr.LocalName == "Id")
                                        rel.Id = xtr.Value;

                                }

                                if (!xtr.LocalName.StartsWith("rsid"))
                                {
                                    xtw.WriteAttributeString(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI, xtr.Value);
                                }
                            }
                            xtr.MoveToElement();
                        }

                        if (isInRel)
                        {
                            isInRel = false;
                            rels.Add(rel);
                        }

                        if (xtr.LocalName.Equals("p"))
                        {
                            xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, _paraId.ToString());
                            xtw.WriteAttributeString(NS_PREFIX, "s", PACKAGE_NS, _sectPrId.ToString());
                            _paraId++;
                        }
                        else if (xtr.LocalName.Equals("sectPr"))
                        {
                            xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, _sectPrId.ToString());
                            _sectPrId++;
                        }

                        if (xtr.IsEmptyElement)
                        {
                            xtw.WriteEndElement();
                        }
                        break;
                    case XmlNodeType.EndElement:
                        xtw.WriteEndElement();
                        break;
                    case XmlNodeType.EntityReference:
                        xtw.WriteEntityRef(xtr.Name);
                        break;
                    case XmlNodeType.ProcessingInstruction:
                        xtw.WriteProcessingInstruction(xtr.Name, xtr.Value);
                        break;
                    case XmlNodeType.SignificantWhitespace:
                        xtw.WriteWhitespace(xtr.Value);
                        break;
                    case XmlNodeType.Text:
                        xtw.WriteString(xtr.Value);
                        break;
                    case XmlNodeType.Whitespace:
                        xtw.WriteWhitespace(xtr.Value);
                        break;
                    case XmlNodeType.XmlDeclaration:
                        // omit XML declaration
                        break;
                    default:
                        Debug.Assert(false);
                        break;
                }
            }
            return rels;
        }

        protected override List<string> Namespaces
        {
            get
            {
                // define the namespaces that are to be copied
                List<string> namespaces = new List<string>();
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                namespaces.Add("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");

                return namespaces;
            }
        }

    }
}
