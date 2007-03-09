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

namespace CleverAge.OdfConverter.Spreadsheet
{
	/// <summary>
	/// Postprocessor which allows to move sharedstrings into cells.
	/// </summary>
	public class OdfSharedStringsPostProcessor : AbstractPostProcessor
	{
		private Stack sharedStringContext;
		private const string PXSI_NAMESPACE = "urn:cleverage:xmlns:post-processings:shared-strings";
		private bool isInSharedString;
		private bool isString;
		private string stringNumber;
		private Element sharedStringElement;
		private bool isPxsi;
		
		public OdfSharedStringsPostProcessor(XmlWriter nextWriter):base(nextWriter)
		{
			this.sharedStringContext = new Stack();
			this.isInSharedString = false;
			this.isString = false;
			this.stringNumber = "0";
			this.isPxsi = false;
		}
		
		public override void WriteStartElement(string prefix, string localName, string ns)
        {
			//sst element starts a sharedstring 
			if(PXSI_NAMESPACE.Equals(ns) && "sst".Equals(localName))
			{
				this.isInSharedString = true;
			}
			//if element is in a sharedstring than push it on the context instead of processing it
			if(isInSharedString)
			{
				this.sharedStringContext.Push(new Element(prefix,localName,ns));
			}
			//v element is going to be replaced by string content
			else if(PXSI_NAMESPACE.Equals(ns) && "v".Equals(localName))
			{
				this.isString = true;
			}
			else
			{
				this.nextWriter.WriteStartElement(prefix, localName, ns);
			}
        }
		public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
			//don't process pxsi attribute because it only contains namespace pxsi
			if("pxsi".Equals(localName))
			{
				this.isPxsi = true;
			}
			//if attribute is in a sharedstring than push it on the context instead of processing it
			else if(isInSharedString)
			{
				this.sharedStringContext.Push(new Attribute(prefix, localName, ns));
			}
			else
			{
				this.nextWriter.WriteStartAttribute(prefix, localName, ns);
			}
        }
        public override void WriteString(string text)
        {
        	//if string is in a sharedstring then add it as a child to top element from context, or assign it as a value of top attribute
        	if(isInSharedString)
        	{
        		Node parentNode = ((Node)sharedStringContext.Peek());
        		if(parentNode is Element)
        		{
        			Element parentElement = ((Element)parentNode);
        			sharedStringContext.Pop();
        			parentElement.AddChild(text);
        			sharedStringContext.Push(parentElement);
        		}
        		else
        		{
        			Attribute parentAttribute = ((Attribute)parentNode);
        			sharedStringContext.Pop();
        			parentAttribute.Value = text;
        			sharedStringContext.Push(parentAttribute);
        		}
        	}
        	//get string number
        	else if(isString)
        	{
        		this.stringNumber = text;
        	}
        	else
        	{
        		this.nextWriter.WriteString(text);
        	}
        }
        public override void WriteEndAttribute()
        {
        	//don't process pxsi attribute
        	if(isPxsi)
        	{
        		this.isPxsi = false;
        	}
        	//if attribute is in sharedstring than pop it from the context and add it to the top element from the context
        	else if(isInSharedString)
        	{
        		Attribute attribute = ((Attribute)sharedStringContext.Peek());
        		sharedStringContext.Pop();
        		Element parentElement =((Element)sharedStringContext.Peek());
        		sharedStringContext.Pop();
        		parentElement.AddAttribute(attribute);
        		sharedStringContext.Push(parentElement);
        	}
        	else
        	{
        		this.nextWriter.WriteEndAttribute();
        	}
        }
		public override void WriteEndElement()
        {
			if(isInSharedString)
			{
				Element element = ((Element)sharedStringContext.Peek());
				sharedStringContext.Pop();
				//ending of sst element ends a sharedstring
				if(PXSI_NAMESPACE.Equals(element.Ns) && "sst".Equals(element.Name))
				{
					this.isInSharedString = false;
					this.sharedStringElement = element;
				}
				//if element is in shared string than add it as a child to top element from the context
				else
				{
					Element parentElement = ((Element)sharedStringContext.Peek());
					sharedStringContext.Pop();
					parentElement.AddChild(element);
					sharedStringContext.Push(parentElement);
				}
			}
			else if(isString)
			{
				//putting sharedstring content into cell
				Element stringElement = sharedStringElement.GetChildWithAttribute("si",PXSI_NAMESPACE,"number",PXSI_NAMESPACE,this.stringNumber);
				if(stringElement != null)
				{
					foreach(object node in stringElement.Children)
					{
						if(node is Element)
						{
							Element nodeElement = ((Element)node);
							nodeElement.Write(this.nextWriter);
						}
						else if(node is string)
						{
							string nodeString = ((string)node);
							this.nextWriter.WriteString(nodeString);
						}
					}
				}
				this.isString = false;
			}
			else
			{
				this.nextWriter.WriteEndElement();
			}
        }
	}
}
