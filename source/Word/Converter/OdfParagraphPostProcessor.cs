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
using System.Text;

namespace OdfConverter.Wordprocessing
{
	/// <summary>
	/// Postprocessor to handle paragraphs with too many characters.
	/// </summary>
	public class OdfParagraphPostProcessor : AbstractPostProcessor
	{
		private Stack currentParagraphNode;
        private Stack currentStyleNode;
        private Stack context;
        private StringBuilder paragraphTextBuilder = new StringBuilder();
        private string styleText;
        private bool bIsDoneStyle = false;
        private ArrayList tabStyleName = new ArrayList();

        private const string PCUT_NAMESPACE = "urn:cleverage:xmlns:post-processings:pcut";
        private const string TEXT_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private const string STYLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private const string FO_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        
        //Extension add to the style name for new styles created
        private const string EXTENSION_STYLE = "_CHILD";

		public OdfParagraphPostProcessor(XmlWriter nextWriter):base(nextWriter)
		{
			this.currentParagraphNode = new Stack();
            this.currentStyleNode = new Stack(); 
            this.context = new Stack();
            this.context.Push("root");
		}

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
        	Element currentElement = new Element(prefix,localName,ns);

            //if element is in paragraph or in style, than we push it into context stack
            if (IsInParagraph() || IsInStyle())
            {                
                this.context.Push(currentElement);
            }
            else
            {
                //if it's a paragraph element, than we start pushing nodes into context stack instead of writing them
                if (localName.Equals("p"))
                {
                    this.currentParagraphNode.Push(currentElement);
                    this.context.Push(currentElement);
                }
                else
                {
                    //if it's a style element, than we start pushing nodes into context stack instead of writing them
                    if (localName.Equals("style"))
                    {
                        this.currentStyleNode.Push(currentElement);
                        this.context.Push(currentElement);
                    }
                    else
                    {
                        this.nextWriter.WriteStartElement(prefix, localName, ns);
                    }
                }
            }
        }
        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
        	//if attribute is in paragraph or in style we push it into context stack
            if (IsInParagraph() || IsInStyle())
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
            if (text!="" && (IsInParagraph() || IsInStyle()))
        	{
        		//if text element in paragraph or in style is not in attribute, it is added to paragraphText or styleText
        		if(IsNotInAttribute())
        		{
        			if(IsInParagraph()) this.paragraphTextBuilder.Append(text);
                    else if(IsInStyle()) this.styleText = this.styleText + text;
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
            if (IsInParagraph() || IsInStyle())
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
        			paragraphTextElement.AddChild(this.paragraphTextBuilder.ToString());
        			element.AddChild(paragraphTextElement);
        			//and then WriteParagraph method is used
        			WriteParagraph(element, false);
        			this.paragraphTextBuilder = new StringBuilder();
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
                if (IsInStyle())
                {
                    Element element = (Element)this.context.Peek();
                    this.context.Pop();
                    object rootElement = this.context.Peek();
                    this.context.Push(element);
                    if (element.Name.Equals("style") && IsRoot(rootElement))
                    { 
                        // a new child element, which contains all the text in paragraph is created
                        Element styleTextElement = new Element("style", "style", element.Ns);
                        styleTextElement.AddChild(this.styleText);
                        element.AddChild(styleTextElement);
                        //and then WriteStyle method is used                        
                        WriteStyle(element, true);
                        this.styleText = "";
                        this.currentStyleNode.Pop();
                        this.context.Pop();
                    }                    
                    else
                    {
                        this.context.Pop();
                        Element parentElement = (Element)this.context.Peek();
                        this.context.Pop();
                        parentElement.AddChild(element);
                        this.context.Push(parentElement);
                    }
                }
                else this.nextWriter.WriteEndElement();
        	}
        }
        //method to check if we are in paragraph
        private bool IsInParagraph()
        {
            try
            {
                // prevent exception from being thrown
                if (this.currentParagraphNode.Count > 0)
                {
                    Element element = (Element)this.currentParagraphNode.Peek();
                    if (element.Name.Equals("p"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
        //method to check if we are in style definition
        private bool IsInStyle()
        {
            try
            {
                // prevent exception from being thrown
                if (this.currentStyleNode.Count > 0)
                {
                    Element element = (Element)this.currentStyleNode.Peek();
                    if (element.Name.Equals("style"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
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
        
        private void WriteParagraph(Element element, bool bCutOneTime)
        {
        	Element textParagraphElement = element.GetChild("paragraph",element.Ns);
        	string textChild = textParagraphElement.GetTextChild();
     		string childTextChild = "";
        	ArrayList childTextChildren = new ArrayList();

            //True if the element flag PCUT_NAMESPACE exist in the paragraph (BUG 1583404)
            if (bCheckPCut(element)) CutParagraph(element, bCutOneTime);
            else
            {
                //if text in paragraph has more than 65535 characters, than the paragraph must be cutted
                if (textChild.Length > 65535)
                {
                    Element nextElement = new Element(element.Prefix, element.Name, element.Ns);
                    string cuttedText = textChild.Substring(0, 65535);
                    bool isTextCutted = false;
                    ArrayList children = (ArrayList)element.Children.Clone();
                    //for each child of the paragraph
                    foreach (Object child in children)
                    {
                        //if we are after the 65535th char, we move child to next paragraph
                        if (isTextCutted)
                        {
                            element.RemoveChild(child);
                            nextElement.AddChild(child);
                        }
                        else
                        {
                            //if child is element, we check if it has text children
                            if (child is Element)
                            {
                                Element childElement = (Element)child;
                                ArrayList childChildren = (ArrayList)childElement.Children.Clone();
                                childTextChildren = childElement.GetTextChildren();
                                Element prevChild = childElement.Clone();
                                Element nextChild = new Element(childElement.Prefix, childElement.Name, childElement.Ns);
                                foreach (Attribute attribute in childElement.Attributes)
                                {
                                    nextChild.AddAttribute(attribute);
                                }
                                if (childTextChildren != null)
                                {
                                    //we move through all children of this child
                                    foreach (object childChild in childChildren)
                                    {
                                        //if we are after the 65535th char we move child to next child
                                        if (isTextCutted)
                                        {
                                            prevChild.RemoveChild(childChild);
                                            nextChild.AddChild(childChild);
                                        }
                                        else
                                        {
                                            if (childChild is string)
                                            {
                                                childTextChild = (string)childChild;
                                            }
                                            if (childTextChild.Length > 0 && cuttedText.Length > 0)
                                            {
                                                //if paragraph text is longer than text in this child we remove it from the beginning of the paragraph text
                                                if (childTextChild.Length < cuttedText.Length)
                                                {
                                                    if (cuttedText.Substring(0, childTextChild.Length).Equals(childTextChild))
                                                    {
                                                        cuttedText = cuttedText.Substring(childTextChild.Length);
                                                    }
                                                }
                                                //else we cut the paragraph
                                                else
                                                {
                                                    if (childTextChild.Substring(0, cuttedText.Length).Equals(cuttedText))
                                                    {
                                                        string previousText = childTextChild.Substring(0, cuttedText.Length);
                                                        string nextText = childTextChild.Substring(cuttedText.Length);
                                                        prevChild.Replace(childTextChild, previousText);
                                                        nextChild.AddChild(nextText);
                                                        isTextCutted = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    //we are replacing this child in paragraph with child with cutted text
                                    element.Replace(child, prevChild);
                                    //the rest of the text is added to a new child of a new paragraph element
                                    if (nextChild.HasChild())
                                    {
                                        nextElement.AddChild(nextChild);
                                    }
                                }
                            }
                            //if child is text, we cut it and put the rest into next paragraph
                            else if (child is string)
                            {
                                string childString = (string)child;
                                if (childString.Length > 0 && cuttedText.Length > 0)
                                {
                                    if (childString.Length < cuttedText.Length)
                                    {
                                        if (cuttedText.Substring(0, childString.Length).Equals(childString))
                                        {
                                            cuttedText = cuttedText.Substring(childString.Length);
                                        }
                                    }
                                    else
                                    {
                                        if (childString.Substring(0, cuttedText.Length).Equals(cuttedText))
                                        {
                                            string previousText = childString.Substring(0, cuttedText.Length);
                                            string nextText = childString.Substring(cuttedText.Length);
                                            element.Replace(child, previousText);
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

                    foreach (Attribute attribute in element.Attributes)
                    {
                        nextElement.AddAttribute(attribute);
                    }
                    //remove all created elements with paragraph text, which wasn't removed before
                    RemoveAllTextParagraphChildren(element);
                    element.Write(this.nextWriter);
                    textParagraphElement.Replace(textChild, textChild.Substring(65535));
                    nextElement.AddChild(textParagraphElement);
                    WriteParagraph(nextElement, false);
                }
                else
                {
                    element.RemoveChild(textParagraphElement);
                    RemoveAllTextParagraphChildren(element);
                    element.Write(this.nextWriter);
                }
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
        private void RemoveAllTextStyleChildren(Element element)
        {
            Element mysteryStyleElement = element.GetChild("style", element.Ns);
            while (mysteryStyleElement != null)
            {
                element.RemoveChild(mysteryStyleElement);
                mysteryStyleElement = element.GetChild("style", element.Ns);
            }
        }

        //Check if there is the element flag PCUT_NAMESPACE to cut the paragraph in two parts
        private bool bCheckPCut(Element element)
        {
            ArrayList children = (ArrayList)element.Children.Clone();
            foreach (Object child in children)
            {
                if (child is Element)
                {
                    Element childElement = (Element)child;
                    if (childElement.Ns == PCUT_NAMESPACE) return true;
                }
            }
            return false;
        }

        //Cut the paragraph in two parts with the element flag PCUT_NAMESPACE
        private void CutParagraph(Element element, bool bCutOneTime)
        {            
            Element textParagraphElement = element.GetChild("paragraph", element.Ns);
            Element nextElement = new Element(element.Prefix, element.Name, element.Ns);
            ArrayList children = (ArrayList)element.Children.Clone();
            bool bMoveChild = false;

            foreach (Object child in children)
            {
                if (child is Element)
                {
                    Element childElement = (Element)child;
                    if (childElement.Ns == PCUT_NAMESPACE && !bMoveChild)
                    {
                        bMoveChild = true;
                        element.RemoveChild(child);
                    }
                    else if (bMoveChild)
                    {
                        element.RemoveChild(child);
                        nextElement.AddChild(child);
                    }                    
                }
                else if (bMoveChild)
                {
                    element.RemoveChild(child);
                    nextElement.AddChild(child);
                }
            }

            element.RemoveChild(textParagraphElement);
            foreach (Attribute attribute in element.Attributes)
            {
                if (attribute.Ns == TEXT_NAMESPACE && attribute.Name == "style-name")
                {
                    Attribute newAttribute = new Attribute(attribute.Prefix, attribute.Name, attribute.Ns);
                    if (!bCutOneTime) newAttribute.Value = attribute.Value + EXTENSION_STYLE;
                    else newAttribute.Value = attribute.Value;
                    nextElement.AddAttribute(newAttribute);
                }
                else nextElement.AddAttribute(attribute);
            }
            RemoveAllTextParagraphChildren(element);
            element.Write(this.nextWriter);
            nextElement.AddChild(textParagraphElement);
            WriteParagraph(nextElement, true);
        }
        
        private void WriteStyle(Element element, bool bIsParentStyle)
        {
            ArrayList children = (ArrayList)element.Children.Clone();
            Element textStyleElement = element.GetChild("style", element.Ns);
            Element nextElement = new Element(element.Prefix, element.Name, element.Ns);
            bool bHaveCutNameSpace = false;

            //Check if we have a cut namespace in the style
            foreach (Object child in children)
            {
                if (child is Element)
                {
                    Element childElement = (Element)child;
                    if (childElement.Name == "paragraph-properties")
                    {
                        Attribute AttributeToDelete = new Attribute("pcut", "pcut", PCUT_NAMESPACE);
                        foreach (Attribute attribute in childElement.Attributes)
                        {
                            if (attribute.Ns == PCUT_NAMESPACE)
                            {
                                bHaveCutNameSpace = true;
                                AttributeToDelete = (Attribute)attribute;                          
                            }
                        }
                        if(bHaveCutNameSpace)
                            childElement.RemoveAttribute(AttributeToDelete);
                    }                  
                }
            }

            //Check if style already write
            foreach (Attribute attribute in element.Attributes)
            {
                if (attribute.Ns == STYLE_NAMESPACE && attribute.Name == "name")
                {
                    if(tabStyleName.Contains(attribute.Value)) bIsDoneStyle = true;
                    else bIsDoneStyle = false;
                }
            }

            //If PCUT_NAMESPACE have been found in attributes
            if (bHaveCutNameSpace && !bIsDoneStyle)
            {
               
                foreach (Object child in children)
                {
                    bool bnewChildElement = false;
                    if (child is Element)
                    {                                                
                        Element childElement = (Element)child;
                        if (childElement.Name == "paragraph-properties")
                        {                            
                            Attribute AttributeToDelete = new Attribute("fo", "break-before", FO_NAMESPACE);
                            string valueAttribute = "";
                            foreach (Attribute attribute in childElement.Attributes)
                            {
                                if (attribute.Ns == FO_NAMESPACE && attribute.Name == "break-before")
                                {
                                    valueAttribute = attribute.Value;
                                    AttributeToDelete = (Attribute)attribute;
                                }
                            }
                            if (AttributeToDelete.Value != "")
                                childElement.RemoveAttribute(AttributeToDelete);
                            
                            //New element paragraph-properties
                            Element newChildElement = new Element(childElement.Prefix, childElement.Name, childElement.Ns);
                            
                            //New Attribute for the element paragraph-properties
                            Attribute newAttributeForChildElement = new Attribute("fo", "break-before", FO_NAMESPACE);
                            newAttributeForChildElement.Value = valueAttribute;
                            newChildElement.AddAttribute(newAttributeForChildElement);

                            foreach (Attribute attribute in childElement.Attributes)
                            {
                                newChildElement.AddAttribute(attribute);
                            }
                            nextElement.AddChild(newChildElement);
                            bnewChildElement = true;
                        }
                        if(!bnewChildElement)
                            nextElement.AddChild(child);
                    }                    
                }             

                //Copy all attributes to the next element
                element.RemoveChild(textStyleElement);
                string nameStyle = "";

                foreach (Attribute attribute in element.Attributes)
                {
                    if (attribute.Ns == STYLE_NAMESPACE && attribute.Name == "name")
                    {
                        Attribute newAttribute = new Attribute(attribute.Prefix, attribute.Name, attribute.Ns);
                        nameStyle = attribute.Value;                        
                        //Add style name in array
                        tabStyleName.Add(attribute.Value);
                        newAttribute.Value = attribute.Value + EXTENSION_STYLE;
                        //Add style name in array for the new style child
                        tabStyleName.Add(newAttribute.Value);
                        nextElement.AddAttribute(newAttribute);
                    }
                    else nextElement.AddAttribute(attribute);
                }

                RemoveAllTextStyleChildren(element);
                element.Write(this.nextWriter);
                nextElement.AddChild(textStyleElement);
                WriteStyle(nextElement, false);                
            }
            else
            {
                element.RemoveChild(textStyleElement);
                RemoveAllTextStyleChildren(element);

                //If Style already write and this is not the parent style (call from WriteStyle)
                //OR
                //If Style not write and this is a parent style (call from WriteEndElement)
                if ((bIsDoneStyle && !bIsParentStyle) || (!bIsDoneStyle && bIsParentStyle))
                    element.Write(this.nextWriter);
            }
        }
	}
}
