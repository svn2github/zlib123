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


namespace CleverAge.OdfConverter.OdfConverterLib
{

	/// <summary>
	/// An <c>XmlWriter</c> implementation for OOX sections post processing
	public class OoxSectionsPostProcessor : AbstractPostProcessor
	{
		private const string OOX_MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		private const string OOX_RELS_NS ="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		private const string CA_SECTIONS_NS = "urn:cleverage:xmlns:post-processings:sections";
		
		private Stack stack;
		private PageContext pageContext;
		
		// test if global page properties have been encountered
		// (ie, properties defined in master styles)
		private bool inGlobal = false;
		// test if local page properties have been encountered
		// (ie, an odf section's columns/notes configuration)
		private bool inLocal = false;
		// test if a master page name has been encountered
		private bool inNameAttr = false;
		private bool inXmlnsPsectAttr = false;
		
		
		public OoxSectionsPostProcessor(XmlWriter nextWriter):base(nextWriter)
		{
			this.stack = new Stack();
			this.pageContext = new PageContext(stack, nextWriter);
		}

		public override void WriteString(string text)
		{
			if (this.inLocal || this.inNameAttr)
			{
				this.pageContext.Set(text);
				this.pageContext.InSectionWriteString(text);
			}
			else
			{
				if (this.inGlobal)
				{
					this.pageContext.InSectionWriteString(text);
				}
				else
				{
					if (!this.inXmlnsPsectAttr)
					{
						nextWriter.WriteString(text);
					}
				}
			}
		}

		
		public override void WriteStartElement(string prefix, string localName, string ns)
		{
			stack.Push(new Element(prefix, localName, ns));
			
			if (InGlobal())
			{
				this.inGlobal = true;
			}
			
			if (InLocal())
			{
				this.inLocal = true;
			}
			
			if (this.inGlobal)
			{
				this.pageContext.InPageStartElement(prefix, localName, ns);
			}
			
			if (this.inLocal)
			{
				this.pageContext.InSectionStartElement(prefix, localName, ns);
			}
			
			if (!this.inGlobal && !this.inLocal)
			{
				this.nextWriter.WriteStartElement(prefix, localName, ns);
			}
		}
		
		
		public override void WriteEndElement()
		{
			if (!this.inGlobal && !this.inLocal)
			{
				this.nextWriter.WriteEndElement();
			}
			
			if (this.inGlobal)
			{
				this.pageContext.InPageEndElement();
			}
			
			if (this.inLocal)
			{
				this.pageContext.InSectionEndElement();
			}
			
			if (InGlobal())
			{
				this.inGlobal = false;
			}
			
			if (InLocal())
			{
				this.inLocal = false;
			}

			stack.Pop();
		}

		
		
		public override void WriteStartAttribute(string prefix, string localName, string ns)
		{
			stack.Push(new Attribute(prefix, localName, ns));
			
			if (InNameAttribute())
			{
				this.inNameAttr = true;
			}
			
			if (InXmlnsPsectAttr())
			{
				this.inXmlnsPsectAttr = true;
			}
			
			if (!this.inGlobal && !this.inLocal && !this.inNameAttr && !this.inXmlnsPsectAttr)
			{
				this.nextWriter.WriteStartAttribute(prefix, localName, ns);
			}
			
			if (this.inGlobal || this.inLocal)
			{
				this.pageContext.InSectionStartAttribute(prefix, localName, ns);
			}
		}

		
		
		public override void WriteEndAttribute()
		{
			if (!this.inGlobal && !this.inLocal && !this.inNameAttr && !this.inXmlnsPsectAttr)
			{
				this.nextWriter.WriteEndAttribute();
			}
			
			if (this.inGlobal || this.inLocal)
			{
				this.pageContext.InSectionEndAttribute();
			}
			
			if (InNameAttribute())
			{
				this.inNameAttr = false;
			}
			
			if (InXmlnsPsectAttr())
			{
				this.inXmlnsPsectAttr = false;
			}
			
			stack.Pop();
		}

		
		
		
		private bool InLocal()
		{
			Node n = (Node) stack.Peek();
			return n.Name.Equals("sectPr") && n.Ns.Equals(OOX_MAIN_NS);
		}
		
		private bool InGlobal()
		{
			Node n = (Node) stack.Peek();
			return n.Name.Equals("master-pages") && n.Ns.Equals(CA_SECTIONS_NS);
		}
		
		private bool InNameAttribute()
		{
			Node n = (Node) stack.Peek();
			return n.Name.Equals("master-page-name") && n.Ns.Equals(CA_SECTIONS_NS);
		}
		
		private bool InXmlnsPsectAttr()
		{
			Node n = (Node) stack.Peek();
			return n.Name.Equals("psect") && n.Prefix.Equals("xmlns");
		}
		
		
		/// <summary>
		/// Keeps track of changes in page styles
		/// </summary>
		protected class PageContext
		{
			private string name;
			private bool skip = false;
			private bool continuous = false;
			private bool pageBreak = false;
			private bool sectionStarts = false;
			private bool sectionEnds = false;
			private Element odfSectPr;
			private Element cont;
			private Element titlePg;
			
			private bool inSectPr = false;
			private Stack sectPrStack;
			private XmlWriter nextWriter;
			
			
			private Stack stack;
			private Hashtable pages;
			
			public PageContext(Stack stack, XmlWriter nextWriter)
			{
				this.stack = stack;
				this.nextWriter = nextWriter;
				this.pages = new Hashtable();
				this.cont = new Element("w", "type", OOX_MAIN_NS);
				this.cont.AddAttribute(new Attribute("w", "val", "continuous", OOX_MAIN_NS));
				this.titlePg = new Element("w", "titlePg", OOX_MAIN_NS);
			}
			
			public Element OdfSectPr
			{
				get { return odfSectPr; }
				set { odfSectPr = value; }
			}
			
			// Create a page style representation based on pre formatted section properties
			public void InPageStartElement(string prefix, string localName, string ns)
			{
				Element e = null;
				if ("master-page".Equals(localName) && CA_SECTIONS_NS.Equals(ns))
				{
					inSectPr = true;
					this.sectPrStack = new Stack();
					e = new Page(prefix, localName, ns);
					sectPrStack.Push(e);
				}
				else
				{
					if (inSectPr) {
						e = new Element(prefix, localName, ns);
						Element parent = (Element) sectPrStack.Peek();
						if (parent != null)
						{
							parent.AddChild(e);
						}
						sectPrStack.Push(e);
					}
				}
			}
			
			// Store pre formatted section properties
			public void InPageEndElement()
			{
				if (inSectPr)
				{
					Element e = (Element) sectPrStack.Pop();
					if ("master-page".Equals(e.Name) && CA_SECTIONS_NS.Equals(e.Ns))
					{
						string masterPageName = e.GetAttributeValue("name", CA_SECTIONS_NS);
						if (pages.Count == 0) 
						{ 	// defaults to the first style
							this.name = masterPageName;
						}
						pages.Add(masterPageName, e);
						inSectPr = false;
					}
				}
			}

			// Store pre formatted section properties
			public void InSectionStartAttribute(string prefix, string localName, string ns)
			{
				if (inSectPr)
				{
					sectPrStack.Push(new Attribute(prefix, localName, ns));
				}
			}
			
			// Store pre formatted section properties
			public void InSectionEndAttribute()
			{
				if (inSectPr) {
					Attribute attr = (Attribute) sectPrStack.Pop();
					if (!attr.Prefix.Equals("xmlns"))
					{
						Node n = (Node) sectPrStack.Peek();
						if (n != null && n is Element)
						{
							Element parent = (Element) n;
							parent.AddAttribute(attr);
						}
					}
				}
			}
			
			// Store pre formatted section properties
			public void InSectionWriteString(string val)
			{
				if (inSectPr) {
					Node n = (Node) sectPrStack.Peek();
					if (n != null && n is Attribute)
					{
						((Attribute) n).Value = val;
					}
				}
			}
			
			// Store preformatted section properties
			public void InSectionStartElement(string prefix, string localName, string ns)
			{
				if ("sectPr".Equals(localName) && OOX_MAIN_NS.Equals(ns))
				{
					inSectPr = true;
					this.sectPrStack = new Stack();
					this.odfSectPr = new Element(prefix, localName, ns);
					sectPrStack.Push(this.OdfSectPr);
				}
				else
				{
					if (inSectPr) {
						Element e = new Element(prefix, localName, ns);
						Element parent = (Element) sectPrStack.Peek();
						if (parent != null)
						{
							parent.AddChild(e);
						}
						sectPrStack.Push(e);
					}
				}
			}
			
			// Produce the oox section and update the page context.
			public void InSectionEndElement()
			{
				if (inSectPr)
				{
					Element e = (Element) sectPrStack.Pop();
					if ("sectPr".Equals(e.Name) && OOX_MAIN_NS.Equals(e.Ns))
					{
						this.Write(nextWriter);
						this.Update();
						inSectPr = false;
					}
				}
			}
			
			
			
			/// <summary>
			/// Gather page context information.
			/// </summary>
			/// <param name="val"></param>
			public void Set(string val) {
				
				Node n = (Node) stack.Peek();
				
				if (n.Ns.Equals(CA_SECTIONS_NS))
				{
					if (n.Name.Equals("master-page-name"))
					{
						this.name = val;
					}
					
					if (n.Name.Equals("page-break"))
					{
						this.pageBreak = true;
						this.skip = false;
						// Determine the next master style name
						if (pages.ContainsKey(this.name))
						{
							Page page = (Page) this.pages[this.name];
							string nextStyle = page.GetAttributeValue("next-style", CA_SECTIONS_NS);
							// a page break with no change in page style => no section needed
							if (nextStyle.Length == 0)
							{
								this.skip = true;
							}
						}
					}
					
					if (n.Name.Equals("section-starts"))
					{
						this.sectionStarts = true;
					}

					if (n.Name.Equals("section-ends"))
					{
						this.sectionEnds = true;
					}
				}
			}
			
			/// <summary>
			/// Update page context.
			/// It may change if a page-break has been encountered and the
			/// current master page name defines a next master style.
			/// </summary>
			public void Update()
			{
				if (this.pageBreak && !this.skip)
				{
					Page page = (Page) this.pages[this.name];
					this.name = page.GetAttributeValue("next-style", CA_SECTIONS_NS);
				}
				
				this.continuous = (this.sectionStarts || this.sectionEnds);
				this.pageBreak = false;
				this.sectionStarts = false;
				this.sectionEnds = false;
				this.skip = false;
				this.odfSectPr = null;
			}
			
			
			/// <summary>
			/// Creates the w:sectPr node
			/// </summary>
			public void Write(XmlWriter nextWriter)
			{
				if (!this.skip && pages.ContainsKey(this.name))
				{
					Page page = (Page) pages[this.name];
					page.Use++;
					
					nextWriter.WriteStartElement("w", "sectPr", OOX_MAIN_NS);
					
					// header / footer reference
					if (page.Use == 1)
					{
						ArrayList hdr = page.GetChildren("headerReference", OOX_MAIN_NS);
						if (hdr.Count > 0) 
						{
							foreach (Element e in hdr)
							{
								e.Write(nextWriter);
							}
						}
						
						ArrayList ftr = page.GetChildren("footerReference", OOX_MAIN_NS);
						if (ftr.Count > 0) 
						{
							foreach (Element e in ftr)
							{
								e.Write(nextWriter);
							}
						}
					}
					
					// footnotes config
					Element footnotePr = null;
					if (this.sectionEnds &&
					    (footnotePr = (Element) this.odfSectPr.GetChild("footnotePr", OOX_MAIN_NS)) != null)
					{
						footnotePr.Write(nextWriter);
					}
					else
					{
						if ((footnotePr = (Element) page.GetChild("footnotePr", OOX_MAIN_NS))!= null)
						{
							footnotePr.Write(nextWriter);
						}
					}
					
					// endnotes config
					Element endnotePr = null;
					if (this.sectionEnds &&
					    (endnotePr = (Element) this.odfSectPr.GetChild("endnotePr", OOX_MAIN_NS)) != null)
					{
						endnotePr.Write(nextWriter);
					}
					else
					{
						if ((endnotePr = (Element) page.GetChild("endnotePr", OOX_MAIN_NS))!= null)
						{
							endnotePr.Write(nextWriter);
						}
					}
					
					// continuity
					if (this.continuous)
					{
						cont.Write(nextWriter);
					}
					
					// pgSz
					Element pgSz = (Element) page.GetChild("pgSz", OOX_MAIN_NS);
					if (pgSz != null)
					{
						pgSz.Write(nextWriter);
					}
					
					// pgMar
					Element pgMar = (Element) page.GetChild("pgMar", OOX_MAIN_NS);
					if (pgMar != null)
					{
						pgMar.Write(nextWriter);
					}
					
					// pgBorders
					Element pgBorders = (Element) page.GetChild("pgBorders", OOX_MAIN_NS);
					if (pgBorders != null)
					{
						pgBorders.Write(nextWriter);
					}
					
					// columns
					Element cols = (Element) page.GetChild("cols", OOX_MAIN_NS);
					if (cols != null)
					{
						cols.Write(nextWriter);
					}
					else
					{
						if (this.sectionEnds &&
						    (cols = (Element) this.odfSectPr.GetChild("cols", OOX_MAIN_NS)) != null)
						{
							cols.Write(nextWriter);
						}
					}
					
					// titlePg
					if (page.GetAttributeValue("name", CA_SECTIONS_NS).Equals("First_20_Page"))
					{
						this.titlePg.Write(nextWriter);
					}
					
					nextWriter.WriteEndElement(); // end sectPr
				}
			}
		}
		
		protected class Page : Element
		{
			private int use;
			
			public int Use {
				get { return use; }
				set { use = value; }
			}
			
			public Page(string prefix, string name, string ns)
				: base(prefix, name, ns) { }
		}
		
		
	}
	
	
	
}
