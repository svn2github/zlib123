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
    /// An <c>XmlWriter</c> implementation for change tracking post processings
    public class OoxChangeTrackingPostProcessor : AbstractPostProcessor
    {

        private const string W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string PCT_NAMESPACE = "urn:cleverage:xmlns:post-processings:change-tracking";

        private string[] PPRCHANGE_CHILDREN = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "textboxTightWrap", "suppressOverlap", "jc", "textDirection", "textAlignment", "outlineLvl" };

        private Stack currentNode;
        private Stack context;
        private Stack regionContext;
        private Stack currentChangedRegion;
        private int currentId = 0;
        private Stack previousParagraph;
        private Element lastRegion;

        /// <summary>
        /// Constructor
        /// </summary>
        public OoxChangeTrackingPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.currentNode = new Stack();
            this.context = new Stack();
            this.context.Push("root");
            this.regionContext = new Stack();
            this.regionContext.Push("root");
            this.currentChangedRegion = new Stack();
            this.previousParagraph = new Stack();
            this.lastRegion = null;
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            this.currentNode.Push(new Element(prefix, localName, ns));
           
            if (IsParagraph(prefix, localName, ns))
            {
                StartParagraph();
            }
            else if (IsStartInsert(prefix, localName, ns))
            {
                StartStartInsert();
            }
            else if (IsEndInsert(prefix, localName, ns))
            {
                StartEndInsert();
            }
            else if (IsInInsert())
            {
                StartElementInInsert(prefix, localName, ns);
            }
            else if (IsInParagraph())
            {
                StartElementInParagraph(prefix, localName, ns);
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteEndElement()
        {
            if (IsInStartInsert())
            {
                EndElementInStartInsert();
            }
            else if (IsInEndInsert())
            {
                EndElementInEndInsert();
            }
            else if (IsInInsert() || IsInRunInsert())
            {
                EndElementInInsert();
            }
            else if (IsInParagraph())
            {
                EndElementInParagraph();
            }
            else
            {
                this.nextWriter.WriteEndElement();
            }

            this.currentNode.Pop();
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            this.currentNode.Push(new Attribute(prefix, localName, ns));

            if (IsInStartInsert() || IsInEndInsert() || IsInParagraph())
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
            if (IsInStartInsert() || IsInEndInsert() || IsInParagraph())
            {
                AddAttributeToElement();
            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }

            this.currentNode.Pop();
        }

        public override void WriteString(string text)
        {
            if (IsInStartInsert() || IsInEndInsert() || IsInParagraph())
            {
                AddStringToNode(text);
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }

        /*
         * General methods
         */

        private void AddAttributeToElement()
        {
            Attribute attribute = (Attribute)this.currentNode.Pop();
            Element element = (Element)this.currentNode.Peek();
            element.AddAttribute(attribute);
            this.currentNode.Push(attribute);
        }

        private void AddStringToNode(string text)
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                Element element = (Element)node;
                element.AddChild(text);
            }
            else
            {
                Attribute attribute = (Attribute)node;
                attribute.Value += text;
            }
        }

        private void AddChildToElement()
        {
            Element child = (Element)this.currentNode.Pop();
            Element parent = (Element)this.currentNode.Peek();
            parent.AddChild(child);
            this.currentNode.Push(child);
        }

        /*
         * Paragraphs
         */

        private bool IsParagraph(string prefix, string localName, string ns)
        {
            return W_NAMESPACE.Equals(ns) && "p".Equals(localName);
        }

        private bool IsInParagraph()
        {
            return this.context.Contains("p") || this.context.Contains("p-with-ins");
        }

        private void StartParagraph()
        {
            if (IsInInsert())
            {
                this.context.Push("p-with-ins");
            }
            else
            {
                this.context.Push("p");
            }
        }

        private void StartElementInParagraph(string prefix, string localName, string ns)
        {
            // do nothing
        }

        private void EndElementInParagraph()
        {
            Element element = (Element)this.currentNode.Peek();
            if (IsParagraph(element.Prefix, element.Name, element.Ns))
            {
                EndParagraph();
            }
            else
            {
                AddChildToElement();
            }
        }

        private void EndParagraph()
        {
            Element paragraph = (Element)this.currentNode.Peek();
            string context = (string)this.context.Pop();
            if ("p-with-ins".Equals(context))
            {
                // we started the paragraph inside an insert region
                // compare para properties with previous para properties and put the previous in pPrChange if they differ
                Element previousParagraph = null;
                if (this.previousParagraph.Count > 0)
                {
                    previousParagraph = (Element)this.previousParagraph.Peek();
                }   
                if (previousParagraph != null && !CompareParagraphProperties(paragraph, previousParagraph))
                {
                    Element pPr = paragraph.GetChild("pPr", W_NAMESPACE);
                    if (pPr == null)
                    {
                        pPr = new Element("w", "pPr", W_NAMESPACE);
                        paragraph.AddChild(pPr);
                    }
                    Element pPrChange = new Element("w", "pPrChange", W_NAMESPACE);
                    Element region = this.lastRegion;
                    if (this.currentChangedRegion.Count > 0)
                    {
                        region = (Element)this.currentChangedRegion.Peek();
                    }
                    pPrChange.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                    pPrChange.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                    pPrChange.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                    Element newPreviousPPr = new Element("w", "pPr", W_NAMESPACE);
                    Element realPreviousPPr = previousParagraph.GetChild("pPr", W_NAMESPACE);
                    if (realPreviousPPr != null)
                    {
                        for (int i = 0; i < PPRCHANGE_CHILDREN.Length; ++i)
                        {
                            Element prop = realPreviousPPr.GetChild(PPRCHANGE_CHILDREN[i], W_NAMESPACE);
                            if (prop != null)
                            {
                                newPreviousPPr.AddChild(prop);
                            }
                        }
                    }
                    pPrChange.AddChild(newPreviousPPr);
                    pPr.AddChild(pPrChange);
                }
            }
            if (this.context.Contains("p") || this.context.Contains("p-with-ins"))
            {
                AddChildToElement();
            }
            else
            {
                paragraph.Write(this.nextWriter);
            }
            if (this.previousParagraph.Count > 0)
            {
                this.previousParagraph.Pop();
            }
            this.previousParagraph.Push(paragraph);
        }

        private void EndParagraphInInsert()
        {
            Element paragraph = (Element)this.currentNode.Peek();
            Element pPr = paragraph.GetChild("pPr", W_NAMESPACE);
            if (pPr == null)
            {
                pPr = new Element("w", "pPr", W_NAMESPACE);
                paragraph.AddChild(pPr);
            }
            Element rPr = pPr.GetChild("rPr", W_NAMESPACE);
            if (rPr == null)
            {
                rPr = new Element("w", "rPr", W_NAMESPACE);
                pPr.AddChild(rPr);
            }
            Element region = (Element)this.currentChangedRegion.Peek();
            Element ins = new Element("w", "ins", W_NAMESPACE);
            ins.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
            ins.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
            ins.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
            rPr.AddFirstChild(ins);
            EndParagraph();
        }

        private bool CompareParagraphProperties(Element p1, Element p2)
        {
            Element pPr1 = p1.GetChild("pPr", W_NAMESPACE);
            Element pPr2 = p2.GetChild("pPr", W_NAMESPACE);
            if (pPr1 == null && pPr2 != null && pPr2.HasChild() || pPr1 != null && (pPr2 == null || !pPr2.HasChild()))
            {
                return false;
            }
            if (pPr1 == null && pPr2 == null)
            {
                return false;
            }
            for (int i = 0; i < PPRCHANGE_CHILDREN.Length; ++i)
            {
                Element prop1 = pPr1.GetChild(PPRCHANGE_CHILDREN[i], W_NAMESPACE);
                Element prop2 = pPr2.GetChild(PPRCHANGE_CHILDREN[i], W_NAMESPACE);
                if (prop1 == null && prop2 != null || prop1 != null && prop2 == null)
                {
                    return false;
                }
                if (prop1 != null && !prop1.IsEqualTo(prop2))
                {
                   return false;
                }
            }
            return true;
        }

        /*
         * StartInsert and EndInsert elements
         */

        private bool IsStartInsert(string prefix, string localName, string ns)
        {
            return PCT_NAMESPACE.Equals(ns) && "start-insert".Equals(localName);
        }

        private bool IsEndInsert(string prefix, string localName, string ns)
        {
            return PCT_NAMESPACE.Equals(ns) && "end-insert".Equals(localName);
        }

        private bool IsInStartInsert()
        {
            return "start-insert".Equals(this.regionContext.Peek());
        }

        private bool IsInEndInsert()
        {
            return "end-insert".Equals(this.regionContext.Peek());
        }

        private void StartStartInsert()
        {
            this.regionContext.Push("start-insert");
        }

        private void StartEndInsert()
        {
            this.regionContext.Push("end-insert");
        }

        private void EndElementInStartInsert()
        {
            // add the region on the stack
            this.currentChangedRegion.Push(this.currentNode.Peek());
            this.regionContext.Pop();
        }

        private void EndElementInEndInsert()
        {
            // remove the region from the regions stack
            Stack tmp = new Stack();
            Element element = (Element)this.currentNode.Peek();
            string id = element.GetAttributeValue("id", PCT_NAMESPACE);
            Element region = (Element)this.currentChangedRegion.Pop();
            while (!id.Equals(region.GetAttributeValue("id", PCT_NAMESPACE)))
            {
                tmp.Push(region);
                region = (Element)this.currentChangedRegion.Pop();
            }
            while (tmp.Count > 0)
            {
                this.currentChangedRegion.Push(tmp.Pop());
            }
            this.regionContext.Pop();
            // save the last region in case we need it before closing a paragraph
            this.lastRegion = region;
        }

        /*
         * Elements within an insert region
         */

        private bool IsInInsert()
        {
            // the second condition has been added in case of a region ending before the run ends.
            return this.currentChangedRegion.Count > 0;
        }

        private bool IsInRunInsert()
        {
             return "r-with-ins".Equals(this.context.Peek());
        }

        private void StartElementInInsert(string prefix, string localName, string ns)
        {
            if (IsInParagraph())
            {
                // normally we should always be in a paragraph
                if (W_NAMESPACE.Equals(ns) && "r".Equals(localName))
                {
                    // run start: we add a w:ins before
                    Element ins = new Element("w", "ins", W_NAMESPACE);
                    ins.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                    Element region = (Element)this.currentChangedRegion.Peek();
                    ins.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                    ins.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                    object current = this.currentNode.Pop();
                    this.currentNode.Push(ins);
                    this.currentNode.Push(current);
                    this.context.Push("r-with-ins");
                }
                StartElementInParagraph(prefix, localName, ns);
            }
            else
            {
                // not sure that can hapens (in a run but not in a paragraph), but in case of...
                if (W_NAMESPACE.Equals(ns) && "r".Equals(localName))
                {
                    this.nextWriter.WriteStartElement("w", "ins", W_NAMESPACE);
                    new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE).Write(this.nextWriter);
                    Element region = (Element)this.currentChangedRegion.Peek();
                    new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE).Write(this.nextWriter);
                    new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE).Write(this.nextWriter);
                    this.context.Push("r-with-ins");
                }
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        private void EndElementInInsert()
        {
            Node currentNode = (Node)this.currentNode.Peek();
            if (IsInParagraph())
            {
                // that is what should normally happen
                if (IsParagraph(currentNode.Prefix, currentNode.Name, currentNode.Ns))
                {
                    EndParagraphInInsert();
                }
                else
                {
                    AddChildToElement();
                    if (W_NAMESPACE.Equals(currentNode.Ns) && "r".Equals(currentNode.Name))
                    {
                        // run end : we close the w:ins
                        this.currentNode.Pop();
                        AddChildToElement();
                        this.context.Pop();
                    }
                }
            }
            else
            {
                // not sure that can happen
                this.nextWriter.WriteEndElement();
                if (W_NAMESPACE.Equals(currentNode.Ns) && "r".Equals(currentNode.Name))
                {
                    // run end : we close the w:ins
                    this.nextWriter.WriteEndElement();
                    this.context.Pop();
                }
            }
        }



    }
}
