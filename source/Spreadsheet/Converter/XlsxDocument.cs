using System;
using System.Collections.Generic;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Xml;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Diagnostics;

namespace CleverAge.OdfConverter.Spreadsheet
{
    class XlsxDocument : OoxDocument
    {
        protected int _partId = 0;

        private const string SPREADSHEET_ML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        
        public XlsxDocument(string fileName) : base(fileName)
        {
        }
        
        protected override List<RelationShip> CopyPart(XmlReader xtr, XmlTextWriter xtw, string ns, string partName)
        {
            bool isInRel = false;
            bool extractRels = ns.Equals(RELATIONSHIP_NS);
            bool isCell = false;
            List<RelationShip> rels = new List<RelationShip>();

            int id = 0;
            
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

                        isCell = xtr.LocalName.Equals("c") && xtr.NamespaceURI.Equals(SPREADSHEET_ML_NS);
                        
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

                                string value = xtr.Value;
                                // normalize type ST_OnOff
                                if (value == "on" || value == "true")
                                    value = "1";
                                else if (value == "off" || value == "false")
                                    value = "0";

                                xtw.WriteAttributeString(xtr.Prefix, xtr.LocalName, xtr.NamespaceURI, value);

                                switch (xtr.LocalName)
                                {
                                    case "r":
                                        if (isCell)
                                        {
                                            string coord = GetColId(value).ToString(System.Globalization.CultureInfo.InvariantCulture)
                                                + "|" 
                                                + GetRowId(value).ToString(System.Globalization.CultureInfo.InvariantCulture);

                                            xtw.WriteAttributeString(NS_PREFIX, "p", PACKAGE_NS, coord);
                                            //xtw.WriteAttributeString(NS_PREFIX, "c", PACKAGE_NS, GetColId(value).ToString(System.Globalization.CultureInfo.InvariantCulture));
                                            //xtw.WriteAttributeString(NS_PREFIX, "r", PACKAGE_NS, GetRowId(value).ToString(System.Globalization.CultureInfo.InvariantCulture));
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
                            case "fonts":
                            case "fills":
                            case "borders":
                            case "cellXfs":
                            case "cellStyleXfs":
                            case "cellStyles":
                            case "dxfs":
                            case "sst":
                                id = 0;
                                break;
                            
                            case "font":
                            case "fill":
                            case "border":
                            case "xf":
                            case "cellStyle":
                            case "dxf":
                            case "si": // sharedStrings
                                xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, (id++).ToString(System.Globalization.CultureInfo.InvariantCulture));
                                break;

                            case "worksheet":
                                id = 0;
                                xtw.WriteAttributeString(NS_PREFIX, "part", PACKAGE_NS, _partId.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                break;
                            case "conditionalFormatting":
                                xtw.WriteAttributeString(NS_PREFIX, "part", PACKAGE_NS, _partId.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                xtw.WriteAttributeString(NS_PREFIX, "id", PACKAGE_NS, (id++).ToString(System.Globalization.CultureInfo.InvariantCulture));
                                break;
                            case "col":
                            case "sheetFormatPr":
                            case "mergeCell":
                            case "drawing":
                            case "hyperlink":
                                xtw.WriteAttributeString(NS_PREFIX, "part", PACKAGE_NS, _partId.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                break;                                
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

            _partId++;
            
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
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable");
                namespaces.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
                namespaces.Add("http://schemas.openxmlformats.org/package/2006/content-types");
                namespaces.Add("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                return namespaces;
            }
        }

        protected int GetRowId(string value)
        {
            int result = 0;

            foreach (char c in value)
	        {
                if (c >= '0' && c <= '9')
    		    {
                    result = (10*result) + (c - '0');
                }
	        }
            
            return result;
        }

        protected int GetColId(string value)
        {
            int result = 0;

            foreach (char c in value)
	        {
                if (c >= 'A' && c <= 'Z')
    		    {
                    result = (26*result) + (c - 'A' + 1);
                }
                else
                {
                    break;
                }
	        }
            
            return result;
        }
    }
}
