using System;
using System.Collections.Generic;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Xml;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Diagnostics;

namespace OdfConverter.Wordprocessing
{
    class DocxDocument : OoxDocument
    {
        protected int _paraId = 0;
        protected int _sectPrId = 0;
        
        protected int _insideField = 0;
        protected int _fieldId = 0;

        public DocxDocument(string fileName) : base(fileName)
        {
        }
        
        protected override void CopyLevel(ZipReader archive, string relFile, XmlTextWriter xtw, List<string> namespaces)
        {
            // init id counters
            if (relFile.Equals("_rels/.rels"))
            {
                _paraId = 0;
                _sectPrId = 0;
            }
            base.CopyLevel(archive, relFile, xtw, namespaces);
        }

        protected override List<RelationShip> CopyPart(XmlReader xtr, XmlTextWriter xtw, string ns, string partName, ZipReader archive)
        {
            bool isInRel = false;
            bool extractRels = ns.Equals(RELATIONSHIP_NS);
            bool fieldBegin = false;
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

                        if (!xtr.LocalName.Equals("proofErr"))
                        {
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
                                        string value = xtr.Value;
                                        // normalize type ST_OnOff
                                        if (value == "on" || value == "true")
                                            value = "1";
                                        else if (value == "off" || value == "false")
                                            value = "0";

                                        xtw.WriteAttributeString(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI, value);
                                    }

                                    switch (xtr.LocalName)
                                    {
                                        case "fldCharType":
                                            if (xtr.Value.Equals("begin"))
                                            {
                                                fieldBegin = true;
                                                _insideField++;
                                                _fieldId++;
                                            }
                                            if (xtr.Value.Equals("end"))
                                            {
                                                _insideField--;
                                            }
                                            break;
                                    }
                                }
                                xtr.MoveToElement();
                            }

                            if (isInRel)
                            {
                                isInRel = false;
                                rels.Add(rel);
                            }

                            switch (xtr.LocalName)
                            {
                                case "p":
                                    xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, _paraId.ToString());
                                    xtw.WriteAttributeString(NS_PREFIX, "s", PACKAGE_NS, _sectPrId.ToString());
                                    _paraId++;
                                    break;
                                case "altChunk":
                                case "bookmarkEnd":
                                case "bookmarkStart":
                                case "commentRangeEnd":
                                case "commentRangeStart":
                                case "del":
                                case "ins":
                                case "moveFrom":
                                case "moveFromRangeEnd":
                                case "moveFromRangeStart":
                                case "moveToRangeEnd":
                                case "moveToRangeStart":
                                case "oMath":
                                case "oMathPara":
                                case "permEnd":
                                case "permStart":
                                case "proofErr":
                                case "sdt":
                                case "tbl":
                                    // write the id of the following w:p
                                    xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, _paraId.ToString());

                                    // write the id of the current section
                                    xtw.WriteAttributeString(NS_PREFIX, "s", PACKAGE_NS, _sectPrId.ToString());
                                    break;
                                case "r":
                                    if (_insideField > 0)
                                    {
                                        // add an attribute if we are inside a field definition
                                        xtw.WriteAttributeString(NS_PREFIX, "f", PACKAGE_NS, _insideField.ToString());
                                        xtw.WriteAttributeString(NS_PREFIX, "fid", PACKAGE_NS, _fieldId.ToString());
                                    }
                                    break;
                                case "sectPr":
                                    xtw.WriteAttributeString(NS_PREFIX, "s", PACKAGE_NS, _sectPrId.ToString());
                                    _sectPrId++;
                                    break;
                                case "instrText":
                                    // add markup to the first instrText so that we can trigger field 
                                    // translation later in template InsertField
                                    if (fieldBegin)
                                    {
                                        xtw.WriteAttributeString(NS_PREFIX, "firstInstrText", PACKAGE_NS, "1");
                                        fieldBegin = false;
                                    }
                                    break;
                            }

                            if (xtr.IsEmptyElement)
                            {
                                xtw.WriteEndElement();
                            }
                        }
                        break;
                    case XmlNodeType.EndElement:
                        if (!xtr.LocalName.Equals("proofErr"))
                        {
                            xtw.WriteEndElement();
                        }
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

            _paraId++;
            _sectPrId++;

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
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml");
                namespaces.Add("http://schemas.openxmlformats.org/package/2006/content-types");

                return namespaces;
            }
        }
    }
}
