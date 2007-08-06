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
using CleverAge.OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.Word
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for change tracking post processings
    public class OoxChangeTrackingPostProcessor : AbstractPostProcessor
    {

        private const string W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string PCT_NAMESPACE = "urn:cleverage:xmlns:post-processings:change-tracking";

        private string[] PPRCHANGE_CHILDREN = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "textboxTightWrap", "suppressOverlap", "jc", "textDirection", "textAlignment", "outlineLvl" };
        private string[] DEL_CHILDREN = { "bookmarkEnd", "bookmarkStart", "commentRangeEnd", "commentRangeStart", "customXml", "customXmlDelRangeEnd", "customXmlDelRangeStart", "customXmlInsRangeEnd", "customXmlInsRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveFromRangeStart", "customXmlMoveToRangeEnd", "customXmlMoveToRangeStart", "del", "ins", "moveFrom", "moveFromRangeEnd", "moveFromRangeStart", "moveTo", "moveToRangeEnd", "moveToRangeStart", "permEnd", "permStart", "proofErr", "r", "sdt", "smartTag"};
        private string[] TR_CHILDREN = { "tblPrEx", "trPr", "tc", "customXml", "sdt", "proofErr", "permStart", "permEnd", "bookmarkStart", "bookmarkEnd", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "commentRangeStart", "commentRangeEnd", "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd", "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd", "ins", "del", "moveFrom", "moveTo"};
        private string[] TRPR_CHILDREN = { "cnfStyle", "divid", "gridBefore", "gridAfter", "wBefore", "wAfter", "canSplit", "trHeight", "trHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange" };
        
        private Stack currentNode;
        private Stack context;
        private Stack currentInsertionRegion;
        private Stack currentFormatChangeRegion;
        private Stack currentDeletion;
        private int currentId = 0;
        private Stack previousParagraph;
        private Element lastInsertionRegion;
        private Element lastFormatChangeRegion;
        private Stack dontWrites;

        /// <summary>
        /// Constructor
        /// </summary>
        public OoxChangeTrackingPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.currentNode = new Stack();
            this.context = new Stack();
            this.context.Push("root");
            this.currentInsertionRegion = new Stack();
            this.currentFormatChangeRegion = new Stack();
            this.currentDeletion = new Stack();
            this.previousParagraph = new Stack();
            this.lastInsertionRegion = null;
            this.lastFormatChangeRegion = null;
            this.dontWrites = new Stack();
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            this.currentNode.Push(new Element(prefix, localName, ns));
            if (IsRun())
            {
                StartRun();
            }
            else if (IsTInDeletion())
            {
                StartTInDeletion();
            }
            else if (IsInstrTextInDeletion())
            {
                StartInstrTextInDeletion();
            }
            else if (IsDeletion())
            {
                StartDeletion();
            }
            else if (IsParagraph())
            {
                StartParagraph();
            }
            else if (IsTableRow())
            {
            	StartTableRow();
            }
            else if (IsStartInsert() || IsEndInsert() || IsStartFormatChange() || IsEndFormatChange() || DontWrite())
            {
                // do nothing
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteEndElement()
        {
            if (IsParagraph())
            {
                EndParagraph();
            }
            else if (IsDeletion())
            {
                EndDeletion();
            }
            else if (IsRun())
            {
                EndRun();
            }
            else if (IsTableRow())
            {
            	EndTableRow();
            }
            else if (IsStartInsert())
            {
                EndStartInsert();
            }
            else if (IsEndInsert())
            {
                EndEndInsert();
            }
            else if (IsStartFormatChange())
            {
                EndStartFormatChange();
            }
            else if (IsEndFormatChange())
            {
                EndEndFormatChange();
            }
            else if (DontWrite())
            {
                AddChildToElement();
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

            if (IsStartInsert() || IsEndInsert() || IsStartFormatChange() || IsEndFormatChange() || DontWrite())
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
            if (IsStartInsert() || IsEndInsert() || IsStartFormatChange() || IsEndFormatChange() || DontWrite())
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
            if (IsStartInsert() || IsEndInsert() || IsStartFormatChange() || IsEndFormatChange() || DontWrite())
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

        private bool DontWrite()
        {
            return this.dontWrites.Count > 0;
        }

        /*
         * Paragraphs
         */

        private bool IsParagraph()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return W_NAMESPACE.Equals(node.Ns) && "p".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsParagraph();
                this.currentNode.Push(node);
                return result;
            }
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
            this.dontWrites.Push("p");
        }

        
     
        
        private void EndParagraph()
        {
            string context = (string)this.context.Pop();
            string dontWrite = (string)this.dontWrites.Pop();
            if (IsInInsert())
            {
                EndParagraphInInsert();
            }
            Element paragraph = (Element)this.currentNode.Peek();
            Element deletion = (Element)paragraph.GetChild("deletion", PCT_NAMESPACE);
            ArrayList followingParagraphs = new ArrayList();
            ArrayList remainingChildren = null;
            // deletion is not null if there were more than one paragraph in it
            if (deletion != null)
            {
                // retrieve all elements after deletion (for insertion in the last paragraph)
                remainingChildren = new ArrayList();
                bool beforeDeletion = true;
                foreach (Object child in paragraph.Children)
                {
                    if (!beforeDeletion)
                    {
                        remainingChildren.Add(child);
                    }
                    else if (child.Equals(deletion))
                    {
                        beforeDeletion = false;
                    }
                }
                // remove all elements after deletion
                foreach (Object child in remainingChildren)
                {
                    paragraph.RemoveChild(child);
                }
                paragraph.RemoveChild(deletion);

                // build w:del element for insertion into paragraphs properties
                Element del = new Element("w", "del", W_NAMESPACE);
                del.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                del.AddAttribute(new Attribute("w", "author", deletion.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                del.AddAttribute(new Attribute("w", "date", deletion.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                // retrieve first paragraph properties
                Element pPr = paragraph.GetChild("pPr", W_NAMESPACE);
                if (pPr == null)
                {
                    pPr = new Element("w", "pPr", W_NAMESPACE);
                    paragraph.AddChild(pPr);
                }

                // extract paragraphs from deletion block
                Element p = deletion.GetChild("p", W_NAMESPACE);
                while (p != null)
                {
                    deletion.RemoveChild(p);
                    followingParagraphs.Add(p);

                    // replace each element with w:del
                    ArrayList children = new ArrayList();
                    foreach (object child in p.Children)
                    {
                        if (child is Element)
                        {
                            Element elChild = (Element)child;
                            if (W_NAMESPACE.Equals(elChild.Ns) && "pPr".Equals(elChild.Name))
                            {
                                children.Add(child);
                            }
                            else
                            {
                                children.Add(ReplaceElementWithDel((Element)child, deletion));
                            }
                        }
                        else
                        {
                            children.Add(child);
                        }
                    }
                    p.Children = children;

                    // replace pPr with first paragraph's one
                    Element pPPr = p.GetChild("pPr", W_NAMESPACE);
                    if (pPPr != null)
                    {
                        p.RemoveChild(pPPr);
                    }
                    Element newPPPr = pPr.Clone();
                    p.AddFirstChild(newPPPr);

                    // add pPrChange if needed
                    if (!CompareParagraphProperties(pPPr, newPPPr))
                    {
                        newPPPr.AddChild(BuildPPrChange(pPPr, deletion));
                    }

                    // is it the last paragraph?
                    Element nextP = deletion.GetChild("p", W_NAMESPACE);
                    if (nextP == null)
                    {
                        // last paragraph: add remaining children
                        foreach (Object child in remainingChildren)
                        {
                            p.AddChild(child);
                        }
                    }
                    else
                    {
                        // previous paragraphs
                        // add w:del to paragraph run properties
                        Element pRPr = newPPPr.GetChild("rPr", W_NAMESPACE);
                        if (pRPr == null)
                        {
                            pRPr = new Element("w", "rPr", W_NAMESPACE);
                            newPPPr.AddChild(pRPr);
                        }
                        pRPr.AddChild(del);
                    }
                    p = nextP;
                }

                // add w:del to paragraph run properties
                Element rPr = pPr.GetChild("rPr", W_NAMESPACE);
                if (rPr == null)
                {
                    rPr = new Element("w", "rPr", W_NAMESPACE);
                    pPr.AddChild(rPr);
                }
                rPr.AddChild(del);
            }
            if ("p-with-ins".Equals(context))
            {
                // we started the paragraph inside an insert region
                // compare para properties with previous para properties and put the previous in pPrChange if they differ
                Element previousParagraph = null;
                if (this.previousParagraph.Count > 0)
                {
                    previousParagraph = (Element)this.previousParagraph.Peek();
                }   
                if (previousParagraph != null && !CompareParagraphProperties(paragraph.GetChild("pPr", W_NAMESPACE), previousParagraph.GetChild("pPr", W_NAMESPACE)))
                {
                    Element pPr = paragraph.GetChild("pPr", W_NAMESPACE);
                    if (pPr == null)
                    {
                        pPr = new Element("w", "pPr", W_NAMESPACE);
                        paragraph.AddChild(pPr);
                    }
                    Element region = this.lastInsertionRegion;
                    if (this.currentInsertionRegion.Count > 0)
                    {
                        region = (Element)this.currentInsertionRegion.Peek();
                    }
                    pPr.AddChild(BuildPPrChange(previousParagraph.GetChild("pPr", W_NAMESPACE), region));
                }
            }
            if (DontWrite())
            {
                AddChildToElement();
            }
            else
            {
                paragraph.Write(this.nextWriter);
            }
            // add other paragraphs (deletion)
            foreach (Element p in followingParagraphs)
            {
                this.currentNode.Pop();
                this.currentNode.Push(p);
                this.dontWrites.Push(dontWrite);
                this.context.Push(context);
                EndParagraph();
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
            Element region = (Element)this.currentInsertionRegion.Peek();
            Element ins = new Element("w", "ins", W_NAMESPACE);
            ins.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
            ins.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
            ins.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
            rPr.AddFirstChild(ins);
        }

        private bool CompareParagraphProperties(Element pPr1, Element pPr2)
        {
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

        private Element BuildPPrChange(Element pPr, Element region)
        {
            Element pPrChange = new Element("w", "pPrChange", W_NAMESPACE);
            pPrChange.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
            pPrChange.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
            pPrChange.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
            Element newPPr = new Element("w", "pPr", W_NAMESPACE);
            if (pPr != null)
            {
                for (int i = 0; i < PPRCHANGE_CHILDREN.Length; ++i)
                {
                    Element prop = pPr.GetChild(PPRCHANGE_CHILDREN[i], W_NAMESPACE);
                    if (prop != null)
                    {
                        newPPr.AddChild(prop);
                    }
                }
            }
            pPrChange.AddChild(newPPr);
            return pPrChange;
        }

        /*
         * Runs
         */

        private bool IsRun()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return W_NAMESPACE.Equals(node.Ns) && "r".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsRun();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsRunInsert()
        {
            return "r-with-ins".Equals(this.context.Peek());
        }

        private void StartRun()
        {
            if (DontWrite())
            {
                if (IsInInsert())
                {
                    // run start: we add a w:ins before
                    Element ins = new Element("w", "ins", W_NAMESPACE);
                    ins.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                    Element region = (Element)this.currentInsertionRegion.Peek();
                    ins.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                    ins.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                    object current = this.currentNode.Pop();
                    this.currentNode.Push(ins);
                    this.currentNode.Push(current);
                    this.context.Push("r-with-ins");
                }
                else
                {
                    this.context.Push("r");
                }
            }
            else
            {
                // not sure that can hapens (in a run but not in a paragraph), but in case of...
                if (IsInInsert())
                {
                    
                    this.nextWriter.WriteStartElement("w", "ins", W_NAMESPACE);
                    new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE).Write(this.nextWriter);
                    Element region = (Element)this.currentInsertionRegion.Peek();
                    new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE).Write(this.nextWriter);
                    new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE).Write(this.nextWriter);
                    this.context.Push("r-with-ins");
                }
                else
                {
                    this.context.Push("r");
                }
                this.nextWriter.WriteStartElement("w", "r", W_NAMESPACE);
            }
        }

        private void EndRun()
        {
            if (IsInFormatChange())
            {
                // format change: add rPrChange
                Element run = (Element)this.currentNode.Peek();
                Element rPr = run.GetChild("rPr", W_NAMESPACE);
                if (rPr == null)
                {
                    rPr = new Element("w", "rPr", W_NAMESPACE);
                    run.AddFirstChild(rPr);
                }
                Element newRPr = rPr.Clone();
                Element rPrChange = new Element("w", "rPrChange", W_NAMESPACE);
                rPrChange.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                Element region = (Element)this.currentFormatChangeRegion.Peek();
                rPrChange.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                rPrChange.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                rPrChange.AddChild(newRPr);
                rPr.AddChild(rPrChange);
            }
            if (DontWrite())
            {
                AddChildToElement();
                if (IsRunInsert())
                {
                    // close the w:ins
                    this.currentNode.Pop();
                    AddChildToElement();
                }
            }
            else
            {
                // not sure that can happen (end run not in a paragraph...)
                this.nextWriter.WriteEndElement();
                if (IsRunInsert())
                {
                    // close the w:ins
                    this.nextWriter.WriteEndElement();
                }
            }
            this.context.Pop();
        }

        /*
         * t and others in deletion
         */

        private bool IsTInDeletion()
        {
            if (!IsInDeletion())
            {
                return false;
            }
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return W_NAMESPACE.Equals(node.Ns) && "t".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsTInDeletion();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsInstrTextInDeletion()
        {
            if (!IsInDeletion())
            {
                return false;
            }
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return W_NAMESPACE.Equals(node.Ns) && "instrText".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsTInDeletion();
                this.currentNode.Push(node);
                return result;
            }
        }

        private void StartTInDeletion()
        {
            Element t = (Element)this.currentNode.Peek();
            t.Name = "delText";
        }

        private void StartInstrTextInDeletion()
        {
            Element t = (Element)this.currentNode.Peek();
            t.Name = "delInstrText";
        }


        /*
         * StartInsert and EndInsert elements
         */

        private bool IsStartInsert()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return PCT_NAMESPACE.Equals(node.Ns) && "start-insert".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsStartInsert();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsEndInsert()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return PCT_NAMESPACE.Equals(node.Ns) && "end-insert".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsEndInsert();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsInInsert()
        {
            return this.currentInsertionRegion.Count > 0;
        }

        private void EndStartInsert()
        {
            // add the region on the stack
            this.currentInsertionRegion.Push(this.currentNode.Peek());
        }

        private void EndEndInsert()
        {
            // remove the region from the regions stack
            Stack tmp = new Stack();
            Element element = (Element)this.currentNode.Peek();
            string id = element.GetAttributeValue("id", PCT_NAMESPACE);
            Element region = (Element)this.currentInsertionRegion.Pop();
            while (!id.Equals(region.GetAttributeValue("id", PCT_NAMESPACE)))
            {
                tmp.Push(region);
                region = (Element)this.currentInsertionRegion.Pop();
            }
            while (tmp.Count > 0)
            {
                this.currentInsertionRegion.Push(tmp.Pop());
            }
            // save the last region in case we need it before closing a paragraph
            this.lastInsertionRegion = region;
        }


        /*
         * StartFormatChange and EndFormatChange elements
         */

        private bool IsStartFormatChange()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return PCT_NAMESPACE.Equals(node.Ns) && "start-format-change".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsStartFormatChange();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsEndFormatChange()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return PCT_NAMESPACE.Equals(node.Ns) && "end-format-change".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsEndFormatChange();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsInFormatChange()
        {
            return this.currentFormatChangeRegion.Count > 0;
        }

        private void EndStartFormatChange()
        {
            // add the region on the stack
            this.currentFormatChangeRegion.Push(this.currentNode.Peek());
        }

        private void EndEndFormatChange()
        {
            // remove the region from the regions stack
            Stack tmp = new Stack();
            Element element = (Element)this.currentNode.Peek();
            string id = element.GetAttributeValue("id", PCT_NAMESPACE);
            Element region = (Element)this.currentFormatChangeRegion.Pop();
            while (!id.Equals(region.GetAttributeValue("id", PCT_NAMESPACE)))
            {
                tmp.Push(region);
                region = (Element)this.currentFormatChangeRegion.Pop();
            }
            while (tmp.Count > 0)
            {
                this.currentFormatChangeRegion.Push(tmp.Pop());
            }
            // save the last region in case we need it before closing a paragraph
            this.lastFormatChangeRegion = region;
        }

        /*
         * Deletions
         */

        private bool IsDeletion()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return PCT_NAMESPACE.Equals(node.Ns) && "deletion".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsDeletion();
                this.currentNode.Push(node);
                return result;
            }
        }

        private bool IsInDeletion()
        {
            return this.context.Contains("del");
        }

        private void StartDeletion()
        {
            this.context.Push("del");
            this.dontWrites.Push("del");
        }

        private void EndDeletion()
        {
            this.context.Pop();
            this.dontWrites.Pop();
            Element deletion = (Element)this.currentNode.Pop();
            Element p = deletion.GetChild("p", W_NAMESPACE);
            // insert elements from first paragraph
            if (p != null)
            {
                // remove the paragraph from the deletion element
                deletion.RemoveChild(p);
                Element parentP = (Element)this.currentNode.Peek();
                foreach (object child in p.Children)
                {
                    if (child is Element)
                    {
                        Element element = (Element)child;
                        if (!"pPr".Equals(element.Name) || !W_NAMESPACE.Equals(element.Ns))
                        {
                            // insert a w:del
                            parentP.AddChild(ReplaceElementWithDel(element, deletion));
                        }
                        /* commented out by jgoffinet to prevent formatting bugs */
                        /*
                        else
                        {
                            // replace paragraph properties from the parent paragraph
                            // (normally they should be the same but sometimes they are not)
                            Element parentPPr = parentP.GetChild("pPr", W_NAMESPACE);
                            if (parentPPr != null)
                            {
                                parentP.RemoveChild(parentPPr);
                            }
                            parentP.AddFirstChild(element);
                        }
                        */
                    }
                }
            }
            this.currentNode.Push(deletion);
            // attach deletion to previous element if not empty
            if (deletion.GetChild("p", W_NAMESPACE) != null)
            {
                AddChildToElement();
            }
        }

        private Element ReplaceElementWithDel(Element element, Element deletion)
        {
            if (W_NAMESPACE.Equals(element.Ns) && Contains(DEL_CHILDREN, element.Name))
            {
                // insert w:del element
                Element del = new Element("w", "del", W_NAMESPACE);
                del.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
                del.AddAttribute(new Attribute("w", "author", deletion.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
                del.AddAttribute(new Attribute("w", "date", deletion.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
                del.AddChild(element);
                return del;
            }
            else
            {
                // recursive call on all children
                ArrayList newChildren = new ArrayList();
                foreach (object child in element.Children)
                {
                    if (child is Element)
                    {
                        Element replaced = ReplaceElementWithDel((Element)child, deletion);
                        newChildren.Add(replaced);
                    }
                    else
                    {
                        newChildren.Add(child);
                    }
                }
                element.Children = newChildren;
                return element;
            }
        }

        private static bool Contains(string[] tab, string s)
        {
            for (int i = 0; i < tab.Length; ++i)
            {
                if (tab[i].Equals(s))
                {
                    return true;
                }
            }
            return false;
        }

        /*
         * Table rows
         */
        private bool IsTableRow()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                return W_NAMESPACE.Equals(node.Ns) && "tr".Equals(node.Name);
            }
            else
            {
                this.currentNode.Pop();
                bool result = IsTableRow();
                this.currentNode.Push(node);
                return result;
            }
        }
        
        private bool IsInTableRow()
        {
        	 return this.context.Contains("tr") || this.context.Contains("tr-with-ins");
        }
        
       	private bool IsTableRowInsert()
        {
            return "tr-with-ins".Equals(this.context.Peek());
        }
        
        private void StartTableRow() 
        {
        	if (IsInInsert())
        	{
        		this.context.Push("tr-with-ins");
        	}
        	else
        	{
        		this.context.Push("tr");
        	}
        	this.dontWrites.Push("tr");
        }
        
        private void EndTableRow()
        {
        	Element tr = (Element) currentNode.Peek();
        		
        	if (IsTableRowInsert()) // only on insertion
        	{
        		Element trPr = tr.GetChild("trPr", W_NAMESPACE);
        		Element region = (Element)this.currentInsertionRegion.Peek();
        		Element ins = new Element("w", "ins", W_NAMESPACE);
        		ins.AddAttribute(new Attribute("w", "id", "" + this.currentId++, W_NAMESPACE));
        		ins.AddAttribute(new Attribute("w", "author", region.GetAttributeValue("creator", PCT_NAMESPACE), W_NAMESPACE));
        		ins.AddAttribute(new Attribute("w", "date", region.GetAttributeValue("date", PCT_NAMESPACE), W_NAMESPACE));
        		if (trPr == null)
        		{
        			trPr = new Element("w", "trPr", W_NAMESPACE);
        			tr.AddChild(trPr);
        			tr.Children = GetOrderedChildren(tr, TR_CHILDREN);
        		}
        		trPr.AddChild(ins);
        		trPr.Children = GetOrderedChildren(trPr, TRPR_CHILDREN);
        		this.context.Pop();
        	}
        	this.dontWrites.Pop();
        	if (DontWrite())
        	{
        		AddChildToElement();
        	}
        	else
        	{
        		tr.Write(nextWriter);
        	}
        }
       
        private ArrayList GetOrderedChildren(Element parent, string [] childrenNames)
        {
        	ArrayList ordered = new ArrayList();
            foreach (string childName in childrenNames)
            {
                Element child = parent.GetChild(childName, W_NAMESPACE);
                if (child != null)
                {
                	ordered.Add(child);
                }
            }
            return ordered;
        }

    }
}
