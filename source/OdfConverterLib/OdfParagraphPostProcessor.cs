﻿/*
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
	/// Postprocessor to handle paragraphs with too many characters.
	/// </summary>
	public class OdfParagraphPostProcessor : AbstractPostProcessor
	{
		private Stack currentParagraphNode;
        private Stack context;
        private string paragraphText;
     
		public OdfParagraphPostProcessor(XmlWriter nextWriter):base(nextWriter)
		{
			this.currentParagraphNode = new Stack();
            this.context = new Stack();
            this.context.Push("root");    
		}

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
        	Element currentElement = new Element(prefix,localName,ns);
        	//if element is in paragraph, than we push it into context stack
        	
        	if(IsInParagraph())
        	{
				this.context.Push(currentElement);
        	}
        	else
        	{
        		//if it's a paragraph element, than we start pushing nodes into context stack instead of writing them
        		if(localName.Equals("p"))
        		{
					this.currentParagraphNode.Push(currentElement);
        			this.context.Push(currentElement);
        		}
        		else
        		{
					this.nextWriter.WriteStartElement(prefix,localName,ns);
        		}
        	}
        }
        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
        	//if attribute is in paragraph we push it into context stack
        	if(IsInParagraph())
        	   {
        	   	Attribute attribute = new Attribute(prefix,localName,ns);
        	   	this.context.Push(attribute);
        	   }
        	   else
        	   {
        	   	this.nextWriter.WriteStartAttribute(prefix,localName,ns);
        	   }
        	}
        public override void WriteString(string text)
        {
        	if(IsInParagraph())
        	{
        		//if text element in paragraph is not in attribute, it is added to paragraphText
        		if(IsNotInAttribute())
        		{
        			this.paragraphText = this.paragraphText + text;
        			Element element = (Element)this.context.Peek();
        			this.context.Pop();
        			element.AddChild(text);
        			this.context.Push(element);
        		}
        		else
        		{
        			Attribute attribute = (Attribute)this.context.Peek();
        			this.context.Pop();
        			attribute.Value = text;
        			this.context.Push(attribute);
        		}
        	}
        	else
        	{
        		this.nextWriter.WriteString(text);
        	}
        }
        public override void WriteEndAttribute()
        {
        	if(IsInParagraph())
        	{
        		Attribute attribute = (Attribute)this.context.Peek();
        		this.context.Pop();
        		Element element = (Element)this.context.Peek();
        		this.context.Pop();
        		element.AddAttribute(attribute);
        		this.context.Push(element);
        	}
        	else
        	{
        		this.nextWriter.WriteEndAttribute();
        	}
        }
        public override void WriteEndElement()
        {
        	if(IsInParagraph())
        	{
        		Element element = (Element)this.context.Peek();
        		this.context.Pop();
        		object rootElement = this.context.Peek();
        		this.context.Push(element);
        		//if it's the end of a main paragraph we write elements from context stack using WriteParagraph method
        		if(element.Name.Equals("p") && IsRoot(rootElement))
        		{
        			// a new child element, which contains all the text in paragraph is created
        			Element paragraphTextElement = new Element("text","paragraph",element.Ns);
        			paragraphTextElement.AddChild(this.paragraphText);
        			element.AddChild(paragraphTextElement);
        			//and then WriteParagraph method is used
        			WriteParagraph(element);
        			this.paragraphText = "";
        			this.currentParagraphNode.Pop();
        			this.context.Pop();
        		}
        		//if it's the end of an element in paragraph, then it is popped from the context stack, and addaed as a child to parentElement
        		else
        		{
        			this.context.Pop();
        			Element parentElement = (Element)this.context.Peek();
        			this.context.Pop();
        			parentElement.AddChild(element);
        			this.context.Push(parentElement);
        		}
        	}
        	else
        	{
        		this.nextWriter.WriteEndElement();
        	}
        }
        //method to check if we are in paragraph
        private bool IsInParagraph()
        {
        	try
        	{
        	Element element = (Element)this.currentParagraphNode.Peek();
        	if(element.Name.Equals("p"))
        	{
        	   	return true;
        	   }
        	   return false;
        	}
        	catch(Exception){
        		return false;
        	}
        }
        //method to check if we are not in attribute
        private bool IsNotInAttribute()
        {
        	try
        	{
        		Object node = this.context.Peek();
        		if(node is Element)
        		{
        			return true;
        		}
        		else
        		{
        			return false;
        		}
        	}
        	catch(Exception)
        	{
        		return false;
        	}
        }
        //method tocheck if element is root(if yes then we are ending paragraph)
        private bool IsRoot(object element)
        {
        	if(element is string)
        	{
        		string text = (string)element;
        		if("root".Equals(text))
        		{
        		   	return true;
        		}
        		return false;
        	}
        	return false;
        }
        
        private void WriteParagraph(Element element)
        {
        	Element textParagraphElement = element.GetChild("paragraph",element.Ns);
        	string textChild = textParagraphElement.GetTextChild();
     		string childTextChild = "";
        	ArrayList childTextChildren = new ArrayList();
        	//if text in paragraph has more than 65535 characters, than the paragraph must be cutted
        	if(textChild.Length > 65535)
        	{
        		Element nextElement = new Element(element.Prefix,element.Name,element.Ns);
        		string cuttedText = textChild.Substring(0,65535);
        		bool isTextCutted = false;
        		ArrayList children = (ArrayList)element.Children.Clone();
        		//for each child of the paragraph
        		foreach(Object child in children)
        		{
        				//if we are after the 65535th char, we move child to next paragraph
        			if(isTextCutted)
        			{
        				element.RemoveChild(child);
        				nextElement.AddChild(child);
        			}
        			else
        			{
        					//if child is element, we check if it has text children
        				if(child is Element)
        				{
        					Element childElement = (Element)child;
        					ArrayList childChildren = (ArrayList)childElement.Children.Clone();
        					childTextChildren = childElement.GetTextChildren();
        					Element prevChild = childElement.Clone();
        					Element nextChild = new Element(childElement.Prefix,childElement.Name,childElement.Ns);
        					foreach(Attribute attribute in childElement.Attributes)
        					{
        						nextChild.AddAttribute(attribute);
        					}
        					if(childTextChildren!=null)
        					{
        						//we move through all children of this child
        						foreach(object childChild in childChildren)
        						{
        							//if we are after the 65535th char we move child to next child
        							if(isTextCutted)
        							{
        								prevChild.RemoveChild(childChild);
        								nextChild.AddChild(childChild);
        							}
        							else
        							{
        								if(childChild is string)
        								{
        									childTextChild = (string)childChild;
        								}
        								if(childTextChild.Length > 0 && cuttedText.Length > 0)
        								{
        									//if paragraph text is longer than text in this child we remove it from the beginning of the paragraph text
        									if(childTextChild.Length < cuttedText.Length)
        									{
        										if(cuttedText.Substring(0,childTextChild.Length).Equals(childTextChild))
        										{
        											cuttedText = cuttedText.Substring(childTextChild.Length);
        										}
        									}
        									//else we cut the paragraph
        									else
        									{
        										if(childTextChild.Substring(0,cuttedText.Length).Equals(cuttedText))
        										{
        											string previousText = childTextChild.Substring(0,cuttedText.Length);
        											string nextText = childTextChild.Substring(cuttedText.Length);
        											prevChild.Replace(childTextChild,previousText);
        											nextChild.AddChild(nextText);
        											isTextCutted = true;
        										}
        									}
        								}
        							}
        						}
        						//we are replacing this child in paragraph with child with cutted text
        						element.Replace(child,prevChild);
        						//the rest of the text is added to a new child of a new paragraph element
        						if(nextChild.HasChild())
        						{
        							nextElement.AddChild(nextChild);
        						}
        					}
        				}
        					//if child is text, we cut it and put the rest into next paragraph
        				else if(child is string)
        				{
        					string childString = (string)child;
        					if(childString.Length > 0 && cuttedText.Length > 0)
        					{
        						if(childString.Length < cuttedText.Length)
        						{
        							if(cuttedText.Substring(0,childString.Length).Equals(childString))
        							{
        								cuttedText = cuttedText.Substring(childString.Length);
        							}
        						}
        						else
        						{
        							if(childString.Substring(0,cuttedText.Length).Equals(cuttedText))
        							{
        								string previousText = childString.Substring(0,cuttedText.Length);
        								string nextText = childString.Substring(cuttedText.Length);
        								element.Replace(child,previousText);
        								nextElement.AddChild(nextText);
        								isTextCutted = true;
        							}
        						}	
        					}
        				}
        			}
        		}
        		//paragraph text element isn't needed any more, and it must be removed
        		element.RemoveChild(textParagraphElement);
        		
        		foreach(Attribute attribute in element.Attributes)
        		{
        			nextElement.AddAttribute(attribute);
        		}
        		//remove all created elements with paragraph text, which wasn't removed before
        		RemoveAllTextParagraphChildren(element);
        		element.Write(this.nextWriter);
        		textParagraphElement.Replace(textChild,textChild.Substring(65535));
        		nextElement.AddChild(textParagraphElement);
        		WriteParagraph(nextElement);
        	}
        	else
        	{
        		element.RemoveChild(textParagraphElement);
        		RemoveAllTextParagraphChildren(element);
        		element.Write(this.nextWriter);
        	}
        }
        private void RemoveAllTextParagraphChildren(Element element)
        {
        	Element mysteryParagraphElement = element.GetChild("paragraph",element.Ns);
        	while(mysteryParagraphElement!=null)
        	{
        		element.RemoveChild(mysteryParagraphElement);
        		mysteryParagraphElement = element.GetChild("paragraph",element.Ns);
        	}	
        }
	}
}
