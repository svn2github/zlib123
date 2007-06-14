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
	/// Postprocessor to remove paragraphs which were processed in Index
	/// </summary>
	public class OdfCheckIfIndexPostProcessor : AbstractPostProcessor
	{
		private int numberOfParagraphs;
		private bool isIndex;
		private int context;
		private Stack context2;
        private int sectionContext;
        private int sectionParagraphs;
		public OdfCheckIfIndexPostProcessor(XmlWriter nextWriter):base(nextWriter)
		{
			this.numberOfParagraphs = 0;
			this.isIndex = false;
			this.context = 0;
			this.context2 = new Stack();
            this.sectionContext = 0;
		}
		public override void WriteStartElement(string prefix, string localName, string ns)
        {
            //field this.sectionContext is increased each time when we start an element in section and decreased when we end an element in section,
            //so when it's value is more than 0, the current element must be in section 
            if (IsSection(localName))
            {
                this.sectionContext = 1;
            }
            else if (this.sectionContext > 0)
            {
                this.sectionContext++;
                if (IsParagraph(localName))
                {
                    this.sectionParagraphs++;
                }
            }
			if(IsIndex(localName))
			{
				this.isIndex = true;
				this.nextWriter.WriteStartElement(prefix,localName,ns);
                if (IsAlphabetical(localName))
                {
                    this.numberOfParagraphs++;
                    //we increse context only if there are no paragraphs between beginning of section and beginning of alphabetical index
                    if (!(this.sectionParagraphs > 0))
                    {
                        this.context++;
                    }
                }
			}
			else
			{
				//If isIndex variable is true, than we are in index and we increase numberOfParagraphs variable
				if(this.isIndex)
				{
					if(IsParagraph(localName))
					{
						this.numberOfParagraphs++;
					}
					this.context++;
					this.nextWriter.WriteStartElement(prefix,localName,ns);
				}
				else
				{
					//don't process paragraphs which should be in index, but they are not after conversion
					if(this.numberOfParagraphs > 0)
					{
						Element element = new Element(prefix,localName,ns);
						this.context2.Push(element);
					}
					else
					{
			    		this.nextWriter.WriteStartElement(prefix,localName,ns);
					}
				
				}
			}
        }
		public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
			//don't process paragraphs which should be in index, but they are not after conversion
			if(this.numberOfParagraphs > 0 && !(this.isIndex))
			{
				
			}
			else
			{
				this.nextWriter.WriteStartAttribute(prefix,localName,ns);
			}
        }
        public override void WriteString(string text)
        {
        	//don't process paragraphs which should be in index, but they are not after conversion
			if(this.numberOfParagraphs > 0 && !(this.isIndex))
        	{
				
			}
			else
			{
				this.nextWriter.WriteString(text);
			}
        }
        public override void WriteEndAttribute()
        {
			//don't process paragraphs which should be in index, but they are not after conversion
			if(this.numberOfParagraphs > 0 && !(this.isIndex))
        	{
				
			}
			else
			{
				this.nextWriter.WriteEndAttribute();
			}
        }
		public override void WriteEndElement()
        {
            //we decrease this.sectionContext field when we end an element in section
            if (this.sectionContext > 0)
            {
                this.sectionContext--;
            }
            else if (this.sectionParagraphs > 0)
            {
                this.sectionParagraphs = 0;
            }
        	if(this.context > 0)
        	{
				this.context--;
        		this.nextWriter.WriteEndElement();
        	}
			else
			{
				if(this.isIndex)
				{
        			this.isIndex = false;
        			this.nextWriter.WriteEndElement();
        			this.numberOfParagraphs--;
				}
				else
				{
					//we decrease numberOfParagraphs variable when we are processing unnecessary paragraphs
					if(this.numberOfParagraphs > 0)
					{
						Element element = (Element)this.context2.Peek();
						if(IsParagraph(element.Name)){
							this.numberOfParagraphs--;
						}
						this.context2.Pop();
					}
					else
					{
						
        				this.nextWriter.WriteEndElement();
        			}
				}
        	}
        }
		//method to check if element starts some kind of index(TOC or bibliography)
		public bool IsIndex(string elementName)
		{
			if(elementName.Equals("table-of-content") || elementName.Equals("bibliography") || elementName.Equals("table-index") || elementName.Equals("alphabetical-index"))
			{
				return true;
			}
			return false;
		}
		
		public bool IsAlphabetical(string elementName)
		{
			if(elementName.Equals("alphabetical-index"))
			   {
			   	return true;
			   }
			   return false;
		}
		//method to check if element is paragraph
		public bool IsParagraph(string elementName)
		{
        	if(elementName.Equals("p"))
        	{
        	   	return true;
        	}
        	return false;
        }
        public bool IsSection(string elementName)
        {
            if (elementName.Equals("section"))
            {
                return true;
            }
            return false;
        } 	
	}
}

