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
using System.Text;
using System.Xml;
using System.Collections;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for paragraphs post processings
    public class OoxParagraphsPostProcessor : AbstractPostProcessor
    {

        private const string W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string DROPCAP_NAMESPACE = "urn:cleverage:xmlns:post-processings:dropcap";
 		private string[] PARAGRAPH_PROPERTIES = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "textboxTightWrap", "suppressOverlap", "jc", "textDirection", "textAlignment", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"};

   		private Stack currentNode;
   		private Stack store;
      	private Element framePr;
   		
        /// <summary>
        /// Constructor
        /// </summary>
        public OoxParagraphsPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
        	this.currentNode = new Stack();
        	this.store = new Stack();
        	this.framePr = new Element("w", "framePr", W_NAMESPACE);
        	this.framePr.AddAttribute(new Attribute("w", "dropCap", "drop", W_NAMESPACE));
        	this.framePr.AddAttribute(new Attribute("w", "vAnchor", "text", W_NAMESPACE));
        	this.framePr.AddAttribute(new Attribute("w", "hAnchor", "text", W_NAMESPACE));
        	this.framePr.AddAttribute(new Attribute("w", "wrap", "around", W_NAMESPACE));
        }

        
        public override void WriteStartElement(string prefix, string localName, string ns)
        {
        	Element e = null;
        	if (W_NAMESPACE.Equals(ns) && "p".Equals(localName))
        	{
        		e = new Paragraph(prefix, localName, ns);
        	}
        	else 
        	{
        		e = new Element(prefix, localName, ns);
        	}
        	
            this.currentNode.Push(e);
            
            if (InParagraph())
            {
            	StartStoreElement();
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
        		WriteStoredParagraph();
        	}
        	if (InParagraph()) 
        	{
        		EndStoreElement();
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

            if (InParagraph())
            {
            	StartStoreAttribute();
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        
        public override void WriteEndAttribute()
        {
        	if (InParagraph())
            {
                EndStoreAttribute();
            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }
            this.currentNode.Pop();
        }

        
        public override void WriteString(string text)
        {
        	if (InParagraph())
            {
   				StoreString(text);
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }

        /*
         * General methods
         */
         
        public void WriteStoredParagraph()
        {
        	Element e = (Element) this.store.Peek();
        	
        	if (e is Paragraph)
        	{
        		Paragraph p = (Paragraph) e;
        		
        		if (p.IsDroppedCap)
        		{
        			Element p0 = SplitParagraph(p);
      
        			if (this.store.Count < 2) 
        			{
        				p0.Write(nextWriter);
        				p.Write(nextWriter);
        			}
        			else 
        			{
        				Element parent = GetParent(p, this.store);
        				if (parent != null)
        				{
        					parent.AddChild(p0);
        				}
        			}
        		}
        		else
        		{
        			if (this.store.Count < 2)
        			{
        				p.Write(nextWriter);
        			}
        		}
        	}
        	else
        	{
        		if (this.store.Count < 2)
        		{
        			e.Write(nextWriter);
        		}
        	}
        }

   		private void StartStoreElement() 
        {
   			Element element = (Element) this.currentNode.Peek();
   			
   			if (this.store.Count > 0)
   			{
   				Element parent = (Element) this.store.Peek();
   				parent.AddChild(element);
   			}
   			this.store.Push(element);	
        }
        
   		
        private void EndStoreElement() 
        {
        	Element e = (Element) this.store.Pop();
        }
        
        
        private void StartStoreAttribute()
        {
        	Element parent = (Element) store.Peek();
        	Attribute attr = (Attribute) this.currentNode.Peek();
        	if (!DROPCAP_NAMESPACE.Equals(attr.Ns) && !"dropcap".Equals(attr.Name))
        	{
        		parent.AddAttribute(attr);
        		this.store.Push(attr);
        	}
        }
        
        
        private void StoreString(string text) 
        {
        	Node node = (Node) this.currentNode.Peek();

        	if (node is Element)
        	{
        		Element element = (Element) this.store.Peek();
        		element.AddChild(text);
        	}
        	else 
        	{
        		if (!DROPCAP_NAMESPACE.Equals(node.Ns) && !"dropcap".Equals(node.Name))
        		{
        			Attribute attr = (Attribute) store.Peek();
        			attr.Value += text;
        		}
        		else
        		{
        			SetDropCapProperties(text);
        		}
        	}
        }
        
        private void EndStoreAttribute()
        {
        	Attribute attr = (Attribute) this.currentNode.Peek();
        	if (!DROPCAP_NAMESPACE.Equals(attr.Ns) && !"dropcap".Equals(attr.Name))
        	{
        		this.store.Pop();
        	}
        }
 
         
        private bool IsParagraph() 
        {
         	Node node = (Node) this.currentNode.Peek();
         	if (node is Element) 
         	{
         		Element element = (Element) node;
         		if ("p".Equals(element.Name) && W_NAMESPACE.Equals(element.Ns))
         		{
         			return true;
         		}
         	}
         	return false;
         }
         
         
         private bool InParagraph()
         {
         	return IsParagraph() || this.store.Count > 0;
         }
          
         
         private void SetDropCapProperties(string text)
         {
         	Paragraph p = (Paragraph) GetBottom(this.store);
         	if (p.DropCapPr == null)
         	{
         		p.DropCapPr = new DropCapProperties();
         	}
         	
         	Attribute attr = (Attribute) this.currentNode.Peek();
         	
         	if ("lines".Equals(attr.Name))
         	{
         		p.DropCapPr.Lines = text;
         	}
         	
 			if ("word".Equals(attr.Name))
         	{
         		p.DropCapPr.IsWord = bool.Parse(text);
         	}
         	
         	if ("length".Equals(attr.Name))
         	{
         		p.DropCapPr.Length = int.Parse(text);
         	}
         	
         	if ("distance".Equals(attr.Name))
         	{
         		p.DropCapPr.Distance = text;
         	}
         		
         	if ("style-name".Equals(attr.Name))
         	{
         		p.DropCapPr.StyleName += text;
         	}
         }
         
         private object GetBottom(Stack stack) 
         {
         	object res = null;
         	int count = stack.Count;
         	for (int i=0; i<count; i++)
         	{
         		res = stack.Peek();
         	}
         	return res;
         }
         
         
         private Element GetParent(Element e, Stack stack)
         {
         	for (int i=0; i<stack.Count; i++)
         	{
         		Node node = (Node) stack.Peek();
         		if (node is Element)
         		{
         			Element parent = (Element) node;
         			foreach (object child in parent.Children)
         			{
         				if (child == e) // object identity
         				{
         					return parent;
         				}
         			}
         		}
         	}
         	return null;
         }
         
         // TODO : test no framePr exist
         // TODO : test this is not an empty text paragraph
         // Split the paragraph in two : 
         // one with the dropcap text in a framePr, the other with the dropcap removed.
         // param p paragraph to be splitten. Will get its dropcap removed
         // returns a new paragraph with the dropcap in a framePr
         private Element SplitParagraph(Paragraph p)
         {
         	Element p0 = new Element("w", "p", W_NAMESPACE);
         	Element pPr = p.GetChild("pPr", W_NAMESPACE);
         	Element pPr0 = null;
         	
         	if (pPr != null)
         	{
         		pPr0 = pPr.Clone();	
         	}
         	else
         	{
         		pPr0 = new Element("w", "pPr", W_NAMESPACE);
         	}
         	p0.AddChild(pPr0);
         	
         	pPr0.AddChildIfNotExist(new Element("w", "keepNext", W_NAMESPACE));
         	Element framePr0 = framePr.Clone();
         	
         	if (p.DropCapPr.Lines != null && p.DropCapPr.Lines.Length > 0)
         	{
         		framePr0.AddAttribute(new Attribute("w", "lines", p.DropCapPr.Lines, W_NAMESPACE));
         	}
         	if (p.DropCapPr.Distance != null && p.DropCapPr.Distance.Length > 0) 
         	{
         		framePr0.AddAttribute(new Attribute("w", "hSpace", p.DropCapPr.Distance, W_NAMESPACE));
         	}
         	
         	pPr0.AddChildIfNotExist(framePr0);
         	Element textAlignment = new Element("w", "textAlignment", W_NAMESPACE);
         	textAlignment.AddAttribute(new Attribute("w", "val", "baseline", W_NAMESPACE));
      		pPr0.AddChild(textAlignment);
         	Element pPrOrdered = GetOrderedParagraphProperties(pPr0);
         	p0.Replace(pPr0, pPrOrdered);
         	
         	
         	foreach (Element run in p.GetChildren("r", W_NAMESPACE))
         	{
         		Element run0 = null;
         		string dropCap = "";	
         		foreach (Element text in run.GetChildren("t", W_NAMESPACE))
         		{
         			ArrayList delayedRemove = new ArrayList();
         			ArrayList delayedReplace = new ArrayList();
         			
         			foreach (object t in text.Children)
         			{	
         				if (t is string)
         				{
         					string s = (string) t;
         					if (!p.DropCapPr.IsWord)
         					{
         						if (p.DropCapPr.Length > 0)
         						{
         							if (s.Length > p.DropCapPr.Length)
         							{
         								dropCap += s.Substring(0, p.DropCapPr.Length);
         								string newS = s.Substring(p.DropCapPr.Length, s.Length-p.DropCapPr.Length);
         								p.DropCapPr.Length -= s.Length;
         								delayedReplace.Add(new Pair(s, newS));
         							}
         							else
         							{
         								dropCap += s;
         								p.DropCapPr.Length -= s.Length;
         								delayedRemove.Add(s);
         							}
         						}
         					}
         					else
         					{
         						int idx = s.IndexOf(' ');
         						if (idx > -1)
         						{
         							dropCap += s.Substring(0, idx);
         							p.DropCapPr.IsWord = false;
         							if (idx+1 < s.Length)
         							{
         								delayedReplace.Add(new Pair(s, s.Substring(idx, s.Length-idx)));
         							}
         							else
         							{
         								dropCap += s;
         								delayedRemove.Add(s);
         							}
         						}
         						else
         						{
         							dropCap += s;
         							delayedRemove.Add(s);
         						}
         					}
         				}
         			}
         			
         			// clean up
         			foreach (string s in delayedRemove) { text.RemoveChild(s); }
         			foreach (Pair pair in delayedReplace) { text.Replace(pair.A, pair.B); }
         		}		
         		
         		if (dropCap.Length > 0)
         		{
         			run0 = new Element("w", "r", W_NAMESPACE);
         			Element rPr = run.GetChild("rPr", W_NAMESPACE);
         			if (rPr != null)
         			{
         				run0.AddChild(rPr);
         			}
         			Element t0 = new Element("w", "t", W_NAMESPACE);
         			t0.AddAttribute(new Attribute("xml", "space", "preserve", null));
         			t0.AddChild(dropCap);
         			run0.AddChild(t0);
         			p0.AddChild(run0);
         		}
         	}
         	
         	return p0;
         }
         
        private Element GetOrderedParagraphProperties(Element pPr)
        {
            Element newPPr = new Element(pPr);
            foreach (string propName in PARAGRAPH_PROPERTIES)
            {
                Element prop = pPr.GetChild(propName, W_NAMESPACE);
                if (prop != null)
                {
                    newPPr.AddChild(prop);
                }
            }
            return newPPr;
        }
        
        
        protected class Paragraph : Element 
        {
        	private DropCapProperties dropCapProperties;
        	
        	public bool IsDroppedCap {
        		get { return dropCapProperties != null; }
        	}
        	
        	public DropCapProperties DropCapPr {
        		get { return dropCapProperties; }
        		set { dropCapProperties = value; }
        	}
        	
        	public Paragraph(Element e) :
        		base (e.Prefix, e.Name, e.Ns) {}
        	
        	public Paragraph (string prefix, string localName, string ns) :
        		base (prefix, localName, ns) {}
        }
      
        
             
        protected class DropCapProperties
        {
        	private bool isWord;
        	private int length;
        	private string distance;
        	private string lines;
        	private string size;
        	private string styleName;
        	
        	public bool IsWord {
        		get { return isWord; }
        		set { isWord = value; }
        	}
        	
        	public int Length {
        		get { return length; }
        		set { length = value; }
        	}
        	
        	public string Distance {
        		get { return distance; }
        		set { distance = value; }
        	}
        	
        	public string Lines {
        		get { return lines; }
        		set { lines = value; }
        	}
        	
        	public string Size {
        		get { return size; }
        		set { size = value; }
        	}
        	
        	public string StyleName {
        		get { return styleName; }
        		set { styleName = value; }
        	}
        }
        
     
        protected class Pair
        {
        	private object a;
        	private object b;
        	
        	public Pair(object a, object b) 
        	{
        		this.a = a;
        		this.b = b;
        	}
        	
        	public object A {
        		get { return a; }
        	}
        	
        	public object B {
        		get { return b; }
        	}
        }
    }
    
}
