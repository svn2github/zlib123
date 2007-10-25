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

using System.Xml;
using System.Collections;
using System;
using CleverAge.OdfConverter.OdfConverterLib;


namespace CleverAge.OdfConverter.Word
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for OOX sections post processing
    public class OoxSectionsPostProcessor : AbstractPostProcessor
    {
        private const string W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PSECT_NAMESPACE = "urn:cleverage:xmlns:post-processings:sections";

        private Stack currentNode;
        private Stack context;
        private Stack store;

        private Element evenAndOddHeaders;
        private int currentPermId = -1;
        private string name;
        private bool skip = false;
        private bool nextIsContinuous = false;
        private bool nextIsPageBreak = false;
        private bool nextIsNewSection = false;
        private bool nextIsEndSection = false;
        private bool nextIsMasterPage = false;
        private bool nextIsSoftPageBreak = false;
        private Element odfSectPr;
        private Element cont;
        private Element titlePg;
        private Hashtable pages;
        private string startPageNumber = null;

       
        public OoxSectionsPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.currentNode = new Stack();
            this.context = new Stack();
            this.evenAndOddHeaders = new Element("w", "evenAndOddHeaders", W_NAMESPACE);
            this.pages = new Hashtable();
            this.cont = new Element("w", "type", W_NAMESPACE);
            this.cont.AddAttribute(new Attribute("w", "val", "continuous", W_NAMESPACE));
            this.titlePg = new Element("w", "titlePg", W_NAMESPACE);
        }

        private Element OdfSectPr
        {
            get { return odfSectPr; }
            set { odfSectPr = value; }
        }

        private bool EvenAndOddHeaders
        {
            get
            {
                ICollection coll = pages.Values;
                foreach (Page page in coll)
                {
                    if (page.EvenOdd)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        /**
         * Overriden callbacks
         */
        public override void WriteString(string text)
        {
            if (InMasterStyles())
            {
                WriteMasterStyles(text);
            }
            else if (InSectPr())
            {
                WriteSectPrAttributes(text);
            }
            else if (InParagraphOrTable())
            {
                WriteParagraphOrTableAttribute(text);
            }
            else if (InEvenAndOddHeaders())
            {
                // nothing to do
            }
            else
            {
                nextWriter.WriteString(text);
            }
        }


        public override void WriteStartElement(string prefix, string localName, string ns)
        {

            this.currentNode.Push(new Element(prefix, localName, ns));

            if (IsMasterStyles())
            {
                StartMasterStyles();
            }
            else if (IsSectPr())
            {
                StartSectPr();
            }
            else if (IsParagraphOrTable())
            {
                StartParagraphOrTable();
            }
            else if (IsEvenAndOddHeaders())
            {
                StartEvenAndOddHeaders();
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteEndElement()
        {
            if (IsPerm())
            {
                EndPerm();
            }
            else if (IsMasterStyles())
            {
                EndMasterStyles();
            }
            else if (IsSectPr())
            {
                EndSectPr();
            }
            else if (IsParagraphOrTable())
            {
                EndParagraphOrTable();
            }
            else if (IsEvenAndOddHeaders())
            {
                EndEvenAndOddHeaders();
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

            if (InMasterStyles())
            {
                StartMasterStylesAttribute();
            }
            else if (InSectPr())
            {
                StartSectPrAttribute();
            }
            else if (InParagraphOrTable())
            {
                StartParagraphOrTableAttribute();
            }
            else if (InEvenAndOddHeaders())
            {
                // nothing to do
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        public override void WriteEndAttribute()
        {
            if (InMasterStyles())
            {
                EndMasterStylesAttribute();
            }
            else if (InSectPr())
            {
                EndSectPrAttribute();
            }
            else if (InParagraphOrTable())
            {
                EndParagraphOrTableAttribute();
            }
            else if (InEvenAndOddHeaders())
            {
                // nothing to do
            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }
            this.currentNode.Pop();
        }

        /**
         * Permission range elements
         */
        private bool IsPerm()
        {
            Element e = (Element)this.currentNode.Peek();
            return ("permStart".Equals(e.Name) || "permEnd".Equals(e.Name))
                && W_NAMESPACE.Equals(e.Ns);
        }

        private void EndPerm()
        {
            Element e = (Element)this.currentNode.Peek();
            if ("permStart".Equals(e.Name))
            {
                this.currentPermId++;
            }
            WritePermId();
            this.nextWriter.WriteEndElement();
        }

        private void WritePermId()
        {
            this.nextWriter.WriteStartAttribute("w", "id", W_NAMESPACE);
            this.nextWriter.WriteString(this.currentPermId.ToString());
            this.nextWriter.WriteEndAttribute();
        }

        /**
         * sectPr elements get skipped or filled with properties 
         * depending on the page context.
         */
        private bool IsSectPr()
        {
            Element e = (Element)this.currentNode.Peek();
            return this.context.Contains("sectPr") ||
                ("sectPr".Equals(e.Name) && W_NAMESPACE.Equals(e.Ns));
        }

        private bool InSectPr()
        {
            return this.context.Contains("sectPr");
        }

        private void StartSectPr()
        {
            Element e = (Element)this.currentNode.Peek();

            if ("sectPr".Equals(e.Name) && W_NAMESPACE.Equals(e.Ns))
            {
                this.context.Push("sectPr");
                this.store = new Stack();
                this.odfSectPr = e;
                this.store.Push(this.OdfSectPr);
            }
            else
            {
                Element parent = (Element)this.store.Peek();
                if (parent != null)
                {
                    parent.AddChild(e);
                }
                this.store.Push(e);
            }
        }

        private void EndSectPr()
        {
            Element e = (Element)this.store.Pop();
            if ("sectPr".Equals(e.Name) && W_NAMESPACE.Equals(e.Ns))
            {
                this.WriteSectPr();
                this.Update();
                this.context.Pop();
            }
        }

        private void StartSectPrAttribute()
        {
            Attribute a = (Attribute)this.currentNode.Peek();
            if (!"psect".Equals(a.Name) && !PSECT_NAMESPACE.Equals(a.Ns))
            {
                Element e = (Element)this.store.Peek();
                e.AddAttribute(a);
                this.store.Push(a);
            }
        }

        private void EndSectPrAttribute()
        {
            Attribute a = (Attribute)this.currentNode.Peek();
            if (!"psect".Equals(a.Name) && !PSECT_NAMESPACE.Equals(a.Ns))
            {
                this.store.Pop();
            }
        }

        private void WriteSectPrAttributes(string text)
        {
            Node node = (Node)this.currentNode.Peek();

            if (PSECT_NAMESPACE.Equals(node.Ns))
            {
                switch (node.Name)
                {
                    case "next-page-break":
                        this.nextIsPageBreak = true;
                        break;
                    case "next-new-section":
                        this.nextIsNewSection = true;
                        break;
                    case "next-end-section":
                        this.nextIsEndSection = true;
                        break;
                    case "next-master-page":
                        this.nextIsMasterPage = true;
                        break;
                    case "next-soft-page-break":
                        this.nextIsSoftPageBreak = true;
                        break;
                }
            }
            else
            {
                if (node is Attribute && !"psect".Equals(node.Name))
                {
                    Attribute a = (Attribute)this.store.Peek();
                    a.Value += text;
                }
            }
        }

        /**
         * Master style definitions get stored in a hashtable
         */
        private bool IsMasterStyles()
        {
            Element e = (Element)this.currentNode.Peek();
            return this.context.Contains("master-pages") ||
                ("master-pages".Equals(e.Name) && PSECT_NAMESPACE.Equals(e.Ns));
        }

        private bool InMasterStyles()
        {
            return this.context.Contains("master-pages");
        }

        private void StartMasterStyles()
        {
            Element e = (Element)this.currentNode.Peek();
            if ("master-pages".Equals(e.Name) && PSECT_NAMESPACE.Equals(e.Ns))
            {
                this.context.Push("master-pages");
            }
            else
            {
                Element elt = null;
                if ("master-page".Equals(e.Name) && PSECT_NAMESPACE.Equals(e.Ns))
                {
                    this.context.Push("master-page");
                    this.store = new Stack();
                    elt = new Page(e.Prefix, e.Name, e.Ns);
                    this.store.Push(elt);
                }
                else
                {
                    if (this.context.Contains("master-page"))
                    {
                        elt = new Element(e);
                        Element parent = (Element)this.store.Peek();
                        if (parent != null)
                        {
                            parent.AddChild(elt);
                        }
                        this.store.Push(elt);
                    }
                }
            }
        }

        private void EndMasterStyles()
        {
            if (this.context.Contains("master-page"))
            {
                Element e = (Element)this.store.Pop();

                if ("master-page".Equals(e.Name) && PSECT_NAMESPACE.Equals(e.Ns))
                {
                    string masterPageName = e.GetAttributeValue("name", PSECT_NAMESPACE);
                    if (this.pages.Count == 0)
                    { 	// defaults to the first style
                        this.name = masterPageName;
                    }
                    // Hack when duplicate name encountered (jgoffinet)
                    if (!this.pages.Contains(masterPageName))
                    {
                        this.pages.Add(masterPageName, e);
                    }
                }
            }
            Element c = (Element)this.currentNode.Peek();
            if (("master-pages".Equals(c.Name) || "master-page".Equals(c.Name)) &&
                PSECT_NAMESPACE.Equals(c.Ns))
            {
                // pop "master-styles"
                this.context.Pop();
            }
        }

        private void StartMasterStylesAttribute()
        {
            if (this.context.Contains("master-page"))
            {
                Attribute a = (Attribute)this.currentNode.Peek();
                Element parent = (Element)this.store.Peek();
                parent.AddAttribute(a);
                this.store.Push(a);

            }
        }

        private void EndMasterStylesAttribute()
        {
            if (this.context.Contains("master-page"))
            {
                this.store.Pop();
            }
        }

        private void WriteMasterStyles(string val)
        {
            if (this.context.Contains("master-page"))
            {
                Node node = (Node)this.store.Peek();
                if (node is Element)
                {
                    Element parent = (Element)node;
                    parent.AddChild(val);
                }
                else
                {
                    Attribute attr = (Attribute)node;
                    attr.Value += val;
                }
            }
        }

        /**
         * Pick up master page names from the text flow.
         */
        private bool IsParagraphOrTable()
        {
            Element e = (Element)this.currentNode.Peek();
            return ("p".Equals(e.Name) || "tbl".Equals(e.Name)) &&
                    W_NAMESPACE.Equals(e.Ns);
        }

        private bool InParagraphOrTable()
        {
            return this.context.Contains("p-or-tbl");
        }

        private void StartParagraphOrTable()
        {
            Element e = (Element)this.currentNode.Peek();
            this.nextWriter.WriteStartElement(e.Prefix, e.Name, e.Ns);
            this.context.Push("p-or-tbl");
        }

        private void EndParagraphOrTable()
        {
            this.nextWriter.WriteEndElement();
            this.context.Pop();
        }

        private void StartParagraphOrTableAttribute()
        {
            Attribute a = (Attribute)this.currentNode.Peek();
            if (!("psect".Equals(a.Name) && "xmlns".Equals(a.Prefix)) && !PSECT_NAMESPACE.Equals(a.Ns))
            {
                this.nextWriter.WriteStartAttribute(a.Prefix, a.Name, a.Ns);
            }
        }

        private void EndParagraphOrTableAttribute()
        {
            Attribute a = (Attribute)this.currentNode.Peek();
            if (!("psect".Equals(a.Name) && "xmlns".Equals(a.Prefix)) && !PSECT_NAMESPACE.Equals(a.Ns))
            {
                this.nextWriter.WriteEndAttribute();
            }
        }

        private void WriteParagraphOrTableAttribute(string text)
        {
            Node node = (Node)this.currentNode.Peek();

            if (PSECT_NAMESPACE.Equals(node.Ns))
            {
                switch (node.Name)
                {
                    case "master-page-name":
                        this.name = text;
                        break;
                    case "page-number":
                        this.startPageNumber = text;
                        break;
                }
            }
            else
            {
                if (!("psect".Equals(node.Name) && "xmlns".Equals(node.Prefix)) && !PSECT_NAMESPACE.Equals(node.Ns))
                {
                    this.nextWriter.WriteString(text);
                }
            }
        }





        /**
         * Change the document settings if we have even and odd headers.
         */
        private bool IsEvenAndOddHeaders()
        {
            Element e = (Element)this.currentNode.Peek();
            return "evenAndOddHeaders".Equals(e.Name) && W_NAMESPACE.Equals(e.Ns);
        }

        private bool InEvenAndOddHeaders()
        {
            return this.context.Contains("evenAndOddHeaders");
        }

        private void StartEvenAndOddHeaders()
        {
            this.context.Push("evenAndOddHeaders");
        }

        private void EndEvenAndOddHeaders()
        {
            if (this.EvenAndOddHeaders)
            {
                evenAndOddHeaders.Write(this.nextWriter);
            }
            this.context.Pop();
        }


        /// <summary>
        /// Update page context.
        /// It may change if a page-break has been encountered and the
        /// current master page name defines a next master style.
        /// </summary>
        public void Update()
        {
            if ((this.nextIsPageBreak || this.nextIsSoftPageBreak) && !this.skip)
            {
                Page page = (Page)this.pages[this.name];
                if (page != null)
                {

                    this.name = page.GetAttributeValue("next-style", PSECT_NAMESPACE);
                }
            }

            if ((!this.nextIsPageBreak || !this.skip) && !this.nextIsSoftPageBreak)
            {
                this.startPageNumber = null;
            }
            
            //clam bugfix #1802267
            this.nextIsContinuous = (this.nextIsNewSection || this.nextIsEndSection || (this.nextIsContinuous && this.skip)) && !this.nextIsPageBreak;
            
            this.nextIsPageBreak = false;
            this.nextIsNewSection = false;
            this.nextIsEndSection = false;
            this.nextIsMasterPage = false;
            this.nextIsSoftPageBreak = false;
            this.skip = false;
            this.odfSectPr = null;
        }



        /// <summary>
        /// Creates the w:sectPr node
        /// </summary>
        private void WriteSectPr()
        {
            if (this.name != null && this.pages.ContainsKey(this.name))
            {
                Page page = (Page)this.pages[this.name];

                // a page break with no change in page style => no section needed
                if ((this.nextIsPageBreak && !this.nextIsMasterPage) || (this.nextIsSoftPageBreak && !this.nextIsMasterPage))
                {
                    string nextStyle = page.GetAttributeValue("next-style", PSECT_NAMESPACE);

                    if (nextStyle.Length == 0)
                    {
                        this.skip = true;
                    }
                }

                if (!this.skip)
                {
                    page.Use++;

                    nextWriter.WriteStartElement("w", "sectPr", W_NAMESPACE);

                    // header / footer reference
                    if (page.Use == 1)
                    {
                        WriteHeaderFooter(page, "headerReference");
                        WriteHeaderFooter(page, "footerReference");

                        // titlePg
                        if (page.FirstDefault)
                        {
                            this.titlePg.Write(nextWriter);
                        }
                    }

                    // footnotes config
                    WriteNotesProperties(page, "footnotePr");
                    WriteNotesProperties(page, "endnotePr");

                    // continuous or even or odd
                    WriteSectionType(page);

                    // page geometry properties
                    WritePageLayout(page);

                   

                    nextWriter.WriteEndElement(); // end sectPr
                }
            }
        }

        // Header and footer references
        private void WriteHeaderFooter(Page page, string localName)
        {
            ArrayList hdr = page.GetChildren(localName, W_NAMESPACE);
            if (hdr.Count > 0)
            {
                foreach (Element e in hdr)
                {
                    if ("even".Equals(e.GetAttributeValue("type", W_NAMESPACE)))
                    {
                        page.EvenOdd = true;                     
                    }
                    if ("first".Equals(e.GetAttributeValue("type", W_NAMESPACE)))
                    {
                        page.FirstDefault = true;
                    }
                    e.Write(nextWriter);
                }
            }
        }

        // Footnotes and endnotes local (section) properties 
        private void WriteNotesProperties(Page page, string localName)
        {
            Element notePr = null;
            if (this.nextIsEndSection &&
                (notePr = (Element)this.odfSectPr.GetChild(localName, W_NAMESPACE)) != null)
            {
                notePr.Write(nextWriter);
            }
            else
            {
                if ((notePr = (Element)page.GetChild(localName, W_NAMESPACE)) != null)
                {
                    notePr.Write(nextWriter);
                }
            }
        }

        // Section type
        private void WriteSectionType(Page page)
        {
            // type
            Element type = (Element)page.GetChild("type", W_NAMESPACE);
            if (type != null)
            {
                type.Write(nextWriter);
            }
            // continuity
            else if (this.nextIsContinuous)
            {
                cont.Write(nextWriter);
            }
        }

        // Page geometry properties
        private void WritePageLayout(Page page)
        {
            // pgSz
            Element pgSz = (Element)page.GetChild("pgSz", W_NAMESPACE);
            if (pgSz != null)
            {
                pgSz.Write(nextWriter);
            }

            // pgMar
            Element pgMar = (Element)page.GetChild("pgMar", W_NAMESPACE);
            if (pgMar != null)
            {
                Element sectMar = (Element)this.odfSectPr.GetChild("pgMar", W_NAMESPACE);
                // if section defines margins, add page and section right/left margin
                if (this.nextIsEndSection && sectMar != null)
                {
                    Element newPgMar = pgMar.Clone();
                    this.OverridePgMargin(newPgMar, sectMar);
                    newPgMar.Write(nextWriter);
                }
                else
                {
                    pgMar.Write(nextWriter);
                }
            }

            // pgBorders
            Element pgBorders = (Element)page.GetChild("pgBorders", W_NAMESPACE);
            if (pgBorders != null)
            {
                pgBorders.Write(nextWriter);
            }

            // lnNumType
            Element lnNumType = (Element)page.GetChild("lnNumType", W_NAMESPACE);
            if (lnNumType != null)
            {
                lnNumType.Write(nextWriter);
            }

            // pgNumType
            Element pgNumType = (Element)page.GetChild("pgNumType", W_NAMESPACE);
            if (pgNumType != null)
            {
                Element pgNumType0 = pgNumType.Clone();
                if (this.startPageNumber != null)
                {
                    pgNumType0.AddAttribute(new Attribute("w", "start", this.startPageNumber, W_NAMESPACE));
                    pgNumType0.Write(nextWriter);
                }
                else if (this.name.Equals("Envelope"))
                {
                    pgNumType0.AddAttribute(new Attribute("w", "start", "0", W_NAMESPACE));
                    pgNumType0.Write(nextWriter);
                }
                else
                {
                    pgNumType0.Write(nextWriter);
                }
            }
            else if (this.name.Equals("Envelope"))
            {
                Element pgNumTypeStart = new Element("w", "pgNumType", W_NAMESPACE);
                pgNumTypeStart.AddAttribute(new Attribute("w", "start", "0", W_NAMESPACE));
                pgNumTypeStart.Write(nextWriter);
            }
            else if (this.startPageNumber != null)
            {
                Element pgNumTypeStart = new Element("w", "pgNumType", W_NAMESPACE);
                pgNumTypeStart.AddAttribute(new Attribute("w", "start", this.startPageNumber, W_NAMESPACE));
                pgNumTypeStart.Write(nextWriter);
            }

            // columns
            Element cols = (Element)this.odfSectPr.GetChild("cols", W_NAMESPACE);
            if (this.nextIsEndSection && cols != null && this.name != "Index")
            {
                cols.Write(nextWriter);
            }
            else
            {
                if ((cols = (Element)page.GetChild("cols", W_NAMESPACE)) != null)
                {
                    cols.Write(nextWriter);
                }
            }
        }

        // Override PgMar element of a page if section defines margin
        private void OverridePgMargin(Element PgMargin, Element SectMargin)
        {
            // right margin
            int rPgVal = 0;
            int rSectVal = 0;
            Attribute rPgMar = PgMargin.GetAttribute("right", W_NAMESPACE);
            Attribute rSectMar = SectMargin.GetAttribute("right", W_NAMESPACE);
            if (rPgMar.Value != null) rPgVal = int.Parse(rPgMar.Value);
            if (rSectMar.Value != null) rSectVal = int.Parse(rSectMar.Value);
            PgMargin.Replace(rPgMar, (new Attribute("w", "right", (rPgVal + rSectVal).ToString(), W_NAMESPACE)));

            // left margin
            int lPgVal = 0;
            int lSectVal = 0;
            Attribute lPgMar = PgMargin.GetAttribute("left", W_NAMESPACE);
            Attribute lSectMar = SectMargin.GetAttribute("left", W_NAMESPACE);
            if (lPgMar.Value != null) lPgVal = int.Parse(lPgMar.Value);
            if (lSectMar.Value != null) lSectVal = int.Parse(lSectMar.Value);
            PgMargin.Replace(lPgMar, (new Attribute("w", "left", (lPgVal + lSectVal).ToString(), W_NAMESPACE)));
        }

        protected class Page : Element
        {
            private int use;
            private bool evenOdd;
            private bool firstDefault;
       
            public int Use
            {
                get { return use; }
                set { use = value; }
            }

            public bool EvenOdd
            {
                get { return evenOdd; }
                set { evenOdd = value; }
            }

            public bool FirstDefault
            {
                get { return firstDefault; }
                set { firstDefault = value; }
            }
          

            public Page(string prefix, string name, string ns)
                : base(prefix, name, ns) { }
        }

    }



}
