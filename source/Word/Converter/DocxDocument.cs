/* 
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
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Xml;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Diagnostics;
using System.Collections;

namespace OdfConverter.Wordprocessing
{
    class DocxDocument : OoxDocument
    {
        protected int _paraId = 0;
        protected int _sectPrId = 0;
        
        protected int _insideField = 0;
        protected int _fieldId = 0;
        protected int _fieldLocked = 0;

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
            bool isInIndex = false;
            List<RelationShip> rels = new List<RelationShip>();
            Stack<string> context = new Stack<string>();

            XmlDocument xmlDocument = new XmlDocument();
            XmlNode bookmarkNode = null;

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

                        if ((xtr.Name.Equals("w:bookmarkEnd") || xtr.Name.Equals("w:bookmarkStart")) 
                            && (xtr.Depth == 1 || (ns.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") && xtr.Depth == 2)))
                        {
                            // opening and closing bookmark tags below w:body (or below w:hdr, w:ftr, w:footnotes etc) will be moved inside the following w:p tag
                            // because in ODF bookmarks can only be opened and closed in paragraph-content
                            bookmarkNode = xmlDocument.CreateElement(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI);
                            if (xtr.HasAttributes)
                            {
                                while (xtr.MoveToNextAttribute())
                                {
                                    XmlAttribute xmlAttribute = xmlDocument.CreateAttribute(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI);
                                    xmlAttribute.Value = xtr.Value;
                                    bookmarkNode.Attributes.Append(xmlAttribute);
                                }
                            }
                        }
                        else if (!xtr.LocalName.Equals("proofErr"))
                        {
                            context.Push(xtr.Name);
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
                                        {
                                            value = "1";
                                        }
                                        else if (value == "off" || value == "false")
                                        {
                                            value = "0";
                                        }

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
                                                fieldBegin = false;
                                                _insideField--;
                                                if (_insideField == 0)
                                                {
                                                    isInIndex = false;
                                                }
                                            }
                                            break;
                                        case "fldLock":
                                            _fieldLocked = (xtr.Value.Equals("on") || xtr.Value.Equals("true") || xtr.Value.Equals("1")) ? 1 : 0;
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

                                    if (bookmarkNode != null)
                                    {
                                        xtw.WriteStartElement(bookmarkNode.Prefix, bookmarkNode.LocalName, bookmarkNode.NamespaceURI);

                                        foreach (XmlAttribute attrib in bookmarkNode.Attributes)
                                        {
                                            xtw.WriteAttributeString(attrib.Prefix, attrib.LocalName, attrib.NamespaceURI, attrib.Value);
                                        }
                                        xtw.WriteEndElement();
                                        bookmarkNode = null;
                                    }

                                    // re-trigger field translation for fields spanning across multiple paragraphs
                                    if (_insideField > 0)
                                    {
                                        fieldBegin = true;
                                    }

                                    if (isInIndex)
                                    {
                                        xtw.WriteAttributeString(NS_PREFIX, "index", PACKAGE_NS, "1");
                                    }

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

                                        // also add the id of the current paragraph so we can get the display value of the 
                                        // field for a specific paragraph easily. This is to support multi-paragraph fields.
                                        xtw.WriteAttributeString(NS_PREFIX, "fpid", PACKAGE_NS, _paraId.ToString());

                                        // also add a marker indicating whether a field is locked or not
                                        xtw.WriteAttributeString(NS_PREFIX, "flocked", PACKAGE_NS, _fieldLocked.ToString());

                                        if (fieldBegin)
                                        {
                                            // add markup to the first run of a field so that we can trigger field 
                                            // translation later in template InsertComplexField
                                            // This is also used to support multi-paragraph fields.
                                            //
                                            xtw.WriteAttributeString(NS_PREFIX, "fStart", PACKAGE_NS, "1");
                                            fieldBegin = false;
                                            _fieldLocked = 0;
                                        }
                                    }
                                    break;
                                case "sectPr":
                                    xtw.WriteAttributeString(NS_PREFIX, "s", PACKAGE_NS, _sectPrId.ToString());
                                    _sectPrId++;
                                    break;
                                case "fldSimple":
                                    if (!xtr.IsEmptyElement)
                                    {
                                        _insideField++;
                                    }
                                    break;
                            }

                            if (xtr.IsEmptyElement)
                            {
                                xtw.WriteEndElement();
                                context.Pop();
                            }
                        }
                        break;
                    case XmlNodeType.EndElement:
                        if (xtr.LocalName.Equals("fldSimple"))
                        {
                            _insideField--;
                        }
                        if (!xtr.LocalName.Equals("proofErr"))
                        {
                            xtw.WriteEndElement();
                            context.Pop();
                        }
                        if (_insideField == 0)
                        {
                            isInIndex = false;
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
                        if (context.Count > 0 && context.Peek().Equals("w:instrText"))
                        {
                            if (xtr.Value.ToUpper().Contains("TOC") || xtr.Value.ToUpper().Contains("BIBLIOGRAPHY") || xtr.Value.ToUpper().Contains("INDEX"))
                            {
                                isInIndex = true;
                            }
                        }
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
