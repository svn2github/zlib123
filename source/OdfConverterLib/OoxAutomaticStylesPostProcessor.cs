/* 
 * Copyright (c) 2006, Clever Age
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
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Xml;
using System.Collections;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for automatic styles post processings
    public class OoxAutomaticStylesPostProcessor : AbstractPostProcessor
    {

        private const string NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private string[] TOGGLE_PROPERTIES = 
            { "b", "bCs", "caps", "emboss", "i", "iCs", "imprint", "outline", "shadow", "smallCaps", "strike", "vanish" };

        private string[] RUN_PROPERTIES = { "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"};
        private string[] PARAGRAPH_PROPERTIES = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "textboxTightWrap", "suppressOverlap", "jc", "textDirection", "textAlignment", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"};

        private Stack currentNode;
        private Stack context;
        private Hashtable cStyles;
        private Hashtable pStyles;

        /// <summary>
        /// Constructor
        /// </summary>
        public OoxAutomaticStylesPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.currentNode = new Stack();
            this.context = new Stack();
            this.context.Push("root");
            this.cStyles = new Hashtable();
            this.pStyles = new Hashtable();
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            this.currentNode.Push(new Element(prefix, localName, ns));

            if (IsStyle(localName, ns))
            {
                StartStyle();
            }
            else if (IsInRun() && IsRunProperties(localName, ns))
            {
                StartRunProperties();
            }
            else if (IsInParagraph() && IsParagraphProperties(localName, ns))
            {
                StartParagraphProperties();
            }
            else if (IsInRunProperties() || IsInParagraphProperties() || IsInStyle())
            {
                // do nothing
            }
            else
            {
                if (IsInRun())
                {
                    StartElementInRun(localName, ns);
                }
                else if (IsRun(localName, ns))
                {
                    StartRun();
                }
                if (IsParagraph(localName, ns))
                {
                    StartParagraph();
                }
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteEndElement()
        {
            if (IsInStyle())
            {
                EndElementInStyle();
            }
            else if (IsInRunProperties())
            {
                EndElementInRunProperties();
            }
            else if (IsInParagraphProperties())
            {
                EndElementInParagraphProperties();
            }
            else
            {
                if (IsInRun())
                {
                    EndElementInRun();
                }
                else if (IsInParagraph())
                {
                    EndElementInParagraph();
                }
                this.nextWriter.WriteEndElement();
            }

            this.currentNode.Pop();
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            this.currentNode.Push(new Attribute(prefix, localName, ns));

            if (IsInRunProperties() || IsInParagraphProperties() || IsInStyle())
            {
                // do nothing
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        public override void WriteEndAttribute()
        {
            if (IsInStyle())
            {
                EndAttributeInStyle();
            }
            else if (IsInRunProperties())
            {
                EndAttributeInRunProperties();
            }
            else if (IsInParagraphProperties())
            {
                EndAttributeInParagraphProperties();
            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }

            this.currentNode.Pop();
        }

        public override void WriteString(string text)
        {
            if (IsInStyle())
            {
                StringInStyle(text);
            }
            else if (IsInRunProperties())
            {
                StringInRunProperties(text);
            }
            else if (IsInParagraphProperties())
            {
                StringInParagraphProperties(text);
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }

        /*
         * Styles declaration
         */

        private bool IsStyle(string localName, string ns)
        {
            return NAMESPACE.Equals(ns) && "style".Equals(localName);
        }

        private bool IsInStyle()
        {
            return "style".Equals(this.context.Peek());
        }

        private void StartStyle()
        {
            this.context.Push("style");
        }

        private void EndElementInStyle()
        {
            Element element = (Element)this.currentNode.Pop();
            if (IsStyle(element.Name, element.Ns))
            {
                // style element : write it if not hidden or not paragraph or character style
                if (IsCharacterStyle(element))
                {
                    string key = element.GetAttributeValue("styleId", NAMESPACE);
                    if (!cStyles.Contains(key))
                    {
                        this.cStyles.Add(key, element);
                    }
                }
                else if (IsParagraphStyle(element))
                {
                    string key = element.GetAttributeValue("styleId", NAMESPACE);
                    if (!pStyles.Contains(key))
                    {
                        this.pStyles.Add(key, element);
                    }
                }
                if (!IsParagraphStyle(element) && !IsCharacterStyle(element) || !IsAutomaticStyle(element))
                {
                    element.Write(this.nextWriter);
                }
                this.context.Pop();
            }
            else
            {
                // child element : add it to its parent
                Element parent = (Element)this.currentNode.Peek();
                parent.AddChild(element);
            }
            this.currentNode.Push(element);
        }

        private void EndAttributeInStyle()
        {
            Attribute attribute = (Attribute)this.currentNode.Pop();
            Element element = (Element)this.currentNode.Peek();
            element.AddAttribute(attribute);
            this.currentNode.Push(attribute);
        }

        private void StringInStyle(string text)
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Attribute)
            {
                Attribute attribute = (Attribute)node;
                attribute.Value = text;
            }
            else if (node is Element)
            {
                Element element = (Element)node;
                element.AddChild(text);
            }
        }

        private bool IsAutomaticStyle(Element style)
        {
            if (style != null)
            {
                return style.GetChild("hidden", NAMESPACE) != null;
            }
            return false;
        }

        private bool IsCharacterStyle(Element style)
        {
            return "character".Equals(style.GetAttributeValue("type", NAMESPACE));
        }

        private bool IsParagraphStyle(Element style)
        {
            return "paragraph".Equals(style.GetAttributeValue("type", NAMESPACE));
        }

        /*
         * Paragraphs
         */

        private string currentParagraphStyleName = null;

        private bool IsParagraph(string localName, string ns)
        {
            return NAMESPACE.Equals(ns) && "p".Equals(localName);
        }

        private bool IsInParagraph()
        {
            return "p".Equals(this.context.Peek());
        }

        private void StartParagraph()
        {
            this.context.Push("p");
        }

        private void EndElementInParagraph()
        {
            Element element = (Element)this.currentNode.Peek();
            if (IsParagraph(element.Name, element.Ns))
            {
                this.currentParagraphStyleName = null;
                this.context.Pop();
            }
        }

        /*
         * Runs
         */

        private bool IsRun(string localName, string ns)
        {
            return NAMESPACE.Equals(ns) && "r".Equals(localName);
        }

        private bool IsInRun()
        {
            return "r".Equals(this.context.Peek()) || "r-with-properties".Equals(this.context.Peek());
        }

        private void StartRun()
        {
            this.context.Push("r");
        }

        private void StartElementInRun(string localName, string ns)
        {
            if (!IsInRunProperties() && "r".Equals(this.context.Peek()))
            {
                if (this.currentParagraphStyleName != null && IsAutomaticStyle((Element)pStyles[this.currentParagraphStyleName]))
                {
                    Element rPr = new Element("w", "rPr", NAMESPACE);
                    AddRunStyleProperties(rPr, this.currentParagraphStyleName, false);
                    DoWriteRunProperties(rPr);
                    this.context.Pop();
                    this.context.Push("r-with-properties");
                }
            }
        }

        private void EndElementInRun()
        {
            Element element = (Element)this.currentNode.Peek();
            if (IsRun(element.Name, element.Ns))
            {
                this.context.Pop();
            }
        }

        /*
         * Run properties
         */

        private bool IsRunProperties(string localName, string ns)
        {
            return NAMESPACE.Equals(ns) && "rPr".Equals(localName);
        }

        private bool IsInRunProperties()
        {
            return "rPr".Equals(this.context.Peek());
        }

        private void StartRunProperties()
        {
            this.context.Push("rPr");
        }

        private void EndElementInRunProperties()
        {
            Element element = (Element)this.currentNode.Pop();
            if (IsRunProperties(element.Name, element.Ns))
            {
                WriteRunProperties(element);
                this.context.Pop();
                this.context.Pop();
                this.context.Push("r-with-properties");
            }
            else
            {
                Element parent = (Element)this.currentNode.Peek();
                parent.AddChild(element);
            }
            this.currentNode.Push(element);
        }

        private void EndAttributeInRunProperties()
        {
            Attribute attribute = (Attribute)this.currentNode.Pop();
            Element element = (Element)this.currentNode.Peek();
            element.AddAttribute(attribute);
            this.currentNode.Push(attribute);
        }

        private void StringInRunProperties(string text)
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Attribute)
            {
                Attribute attribute = (Attribute)node;
                attribute.Value = text;
            }
            else if (node is Element)
            {
                Element element = (Element)node;
                element.AddChild(text);
            }
        }

        private void WriteRunProperties(Element rPr)
        {
            Element rStyle = rPr.GetChild("rStyle", NAMESPACE);
            if (rStyle != null)
            {
                // remove style declaration (will add it later back)
                rPr.RemoveChild(rStyle);
                // add run style properties
                string rStyleName = rStyle.GetAttributeValue("val", NAMESPACE);
                if (rStyleName.Length > 0)
                {
                    AddRunStyleProperties(rPr, rStyleName, true);
                }
            }
            // add paragraph run properties (if automatic)
            if (this.currentParagraphStyleName != null && IsAutomaticStyle((Element)pStyles[this.currentParagraphStyleName]))
            {
                Element style = (Element)pStyles[this.currentParagraphStyleName];
                if (style != null)
                {
                    AddRunProperties(rPr, style);
                }
            }
            // add style name
            if (rStyle != null)
            {
                string rStyleName = rStyle.GetAttributeValue("val", NAMESPACE);
                if (rStyleName.Length > 0)
                {
                    AddStyleDeclaration(rPr, rStyleName, "rStyle");
                }
            }
            // write run properties
            DoWriteRunProperties(rPr);
        }

        private void AddRunStyleProperties(Element rPr, string styleName, bool isCharacterStyle)
        {
            Element style = null;
            if (isCharacterStyle)
            {
                style = (Element)cStyles[styleName];
            }
            else
            {
                style = (Element)pStyles[styleName];
            }

            // to avoid crash when wrong styles applied
            if (style == null)
            {
                return;
            }

            // add all run properties
            AddRunProperties(rPr, style);

            // add parent style properties
            Element basedOn = (Element)style.GetChild("basedOn", NAMESPACE);
            if (basedOn != null)
            {
                string val = basedOn.GetAttributeValue("val", NAMESPACE);
                if (val.Length > 0)
                {
                    AddRunStyleProperties(rPr, val, isCharacterStyle);
                }
            }
        }

        private void AddRunProperties(Element rPr, Element style)
        {
            Element styleRPr = (Element)style.GetChild("rPr", NAMESPACE);
            if (styleRPr != null)
            {
                foreach (Element prop in styleRPr.Children)
                {
                    if (IsAutomaticStyle(style) || isToggle(prop))
                    {
                        rPr.AddChildIfNotExist(prop);
                    }
                }
            }
        }

        private void DoWriteRunProperties(Element rPr)
        {
            this.nextWriter.WriteStartElement(rPr.Prefix, rPr.Name, rPr.Ns);
            foreach (string propName in RUN_PROPERTIES)
            {
                Element prop = rPr.GetChild(propName, NAMESPACE);
                if (prop != null)
                {
                    prop.Write(this.nextWriter);
                }
            }
            this.nextWriter.WriteEndElement();
        }

        private bool isToggle(Element prop)
        {
            for (int i = 0; i < TOGGLE_PROPERTIES.Length; ++i)
            {
                if (TOGGLE_PROPERTIES[i].Equals(prop.Name))
                {
                    return true;
                }
            }
            return false;
        }

        /*
         * Paragraph properties
         */

        private bool IsParagraphProperties(string localName, string ns)
        {
            return NAMESPACE.Equals(ns) && "pPr".Equals(localName);
        }

        private bool IsInParagraphProperties()
        {
            return "pPr".Equals(this.context.Peek());
        }

        private void StartParagraphProperties()
        {
            this.context.Push("pPr");
        }

        private void EndElementInParagraphProperties()
        {
            Element element = (Element)this.currentNode.Pop();
            if (IsParagraphProperties(element.Name, element.Ns))
            {
                WriteParagraphProperties(element);
                this.context.Pop();
            }
            else
            {
                Element parent = (Element)this.currentNode.Peek();
                parent.AddChild(element);
            }
            this.currentNode.Push(element);
        }

        private void EndAttributeInParagraphProperties()
        {
            Attribute attribute = (Attribute)this.currentNode.Pop();
            Element element = (Element)this.currentNode.Peek();
            element.AddAttribute(attribute);
            this.currentNode.Push(attribute);
        }

        private void StringInParagraphProperties(string text)
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Attribute)
            {
                Attribute attribute = (Attribute)node;
                attribute.Value = text;
            }
            else if (node is Element)
            {
                Element element = (Element)node;
                element.AddChild(text);
            }
        }

        private void WriteParagraphProperties(Element pPr)
        {
            Element rPr = pPr.GetChild("rPr", NAMESPACE);
            if (rPr == null)
            {
                rPr = new Element("w", "rPr", NAMESPACE);
            }
            else
            {
                pPr.RemoveChild(rPr);
            }
            Element pStyle = pPr.GetChild("pStyle", NAMESPACE);
            if (pStyle != null)
            {
                // remove style declaration (will be added later back)
                pPr.RemoveChild(pStyle);
                // add paragraph style properties
                string pStyleName = pStyle.GetAttributeValue("val", NAMESPACE);
                if (pStyleName.Length > 0)
                {
                    Element style = (Element)this.pStyles[pStyleName];
                    if (style != null)
                    {
                        AddParagraphProperties(pPr, style);
                        AddRunStyleProperties(rPr, pStyleName, false);
                        this.currentParagraphStyleName = pStyleName;
                        // add style declaration
                        AddStyleDeclaration(pPr, pStyleName, "pStyle");
                    }
                }
            }
            // add run properties
            if (rPr.HasChild())
            {
                pPr.AddChild(rPr);
            }
            // Write paragraph properties
            DoWriteParagraphProperties(pPr);
        }

        private void AddParagraphStyleProperties(Element pPr, string styleName)
        {
            Element style = (Element)pStyles[styleName];
            // to avoid crash when wrong styles applied
            if (style == null)
            {
                return;
            }
            if (IsAutomaticStyle(style))
            {
                // add run properties
                AddParagraphProperties(pPr, style);

                // add parent style properties
                Element basedOn = (Element)style.GetChild("basedOn", NAMESPACE);
                if (basedOn != null)
                {
                    string val = basedOn.GetAttributeValue("val", NAMESPACE);
                    if (val.Length > 0)
                    {
                        AddParagraphStyleProperties(pPr, val);
                    }
                }
            }
        }

        private void AddParagraphProperties(Element pPr, Element style)
        {
            Element stylePPr = (Element)style.GetChild("pPr", NAMESPACE);
            if (stylePPr != null)
            {
                foreach (Element prop in stylePPr.Children)
                {
                    pPr.AddChildIfNotExist(prop);
                }
            }
        }

        private void DoWriteParagraphProperties(Element pPr)
        {
            this.nextWriter.WriteStartElement(pPr.Prefix, pPr.Name, pPr.Ns);
            foreach (string propName in PARAGRAPH_PROPERTIES)
            {
                Element prop = pPr.GetChild(propName, NAMESPACE);
                if (prop != null)
                {
                    if ("rPr".Equals(propName))
                    {
                        DoWriteRunProperties(prop);
                    }
                    else
                    {
                        prop.Write(this.nextWriter);
                    }
                }
            }
            this.nextWriter.WriteEndElement();
        }


        // Adds a style declaration by looking for the first non automatic style among the parents
        private void AddStyleDeclaration(Element element, string styleName, string styleType)
        {
            Element style = null;
            if ("rStyle".Equals(styleType))
            {
                style = (Element)cStyles[styleName];
            }
            else
            {
                style = (Element)pStyles[styleName];
            }

            // to avoid crash when wrong styles applied
            if (style == null)
            {
                return;
            }

            
            if (IsAutomaticStyle(style))
            {
                Element basedOn = (Element)style.GetChild("basedOn", NAMESPACE);
                if (basedOn != null)
                {
                    string val = basedOn.GetAttributeValue("val", NAMESPACE);
                    if (val.Length > 0)
                    {
                        AddStyleDeclaration(element, val, styleType);
                    }
                }
            }
            else
            {
                Element decl = new Element("w", styleType, NAMESPACE);
                Attribute val = new Attribute("w", "val", NAMESPACE);
                val.Value = styleName;
                decl.AddAttribute(val);
                element.AddFirstChild(decl);
            }
        }


    }
}
