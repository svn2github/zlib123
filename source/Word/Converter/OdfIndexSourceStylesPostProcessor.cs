/* 
 * Copyright (c) 2006-2008 Clever Age
 * Copyright (c) 2006-2009 DIaLOGIKa
 *
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
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Xml;
using System.Collections;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Collections.Generic;

namespace OdfConverter.Wordprocessing
{
    /// <summary>
    /// Postprocessor to insert Index Source Styles.
    /// </summary>
    public class OdfIndexSourceStylesPostProcessor : AbstractPostProcessor
    {
        private const string TEXT_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private const string PXSI_NAMESPACE = "urn:cleverage:xmlns:post-processings:source-styles";
        private bool IsIndexSourceStyleProcessed;
        public OdfIndexSourceStylesPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.IsIndexSourceStyleProcessed = false;   
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            if ("index-source-styles".Equals(localName) && PXSI_NAMESPACE.Equals(ns))
            { 
                this.IsIndexSourceStyleProcessed = true;
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }
        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            if (!IsIndexSourceStyleProcessed)
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        // math, dialogika: Bugfix #1841783: Parsing of styles and levels was wrong. 

        public override void WriteString(string text)
        {
            if (IsIndexSourceStyleProcessed)
            {
                bool indexSourceStylesContained = false;
                for (int i = 1; i < 11; i++)
                {
                    if (text.Contains(":" + i))
                    {
                        indexSourceStylesContained =true;
                        break;
                    }
                }

                if (!indexSourceStylesContained)
                {
                    return;
                }

                List<string>[] styles = new List<string>[11];
                string[] stylesWithLevel = text.Split('.');
                try
                {
                    foreach (string s in stylesWithLevel)
                    {
                        if (s != "")
                        {
                            uint level = Convert.ToUInt32(s.Substring(s.IndexOf(':') + 1));
                            string style = s.Substring(0, s.IndexOf(':'));
                            
                            if (styles[level] == null)
                            {
                                styles[level] = new List<string>();
                            }
                            styles[level].Add(style);
                        }
                    }
                }
                catch (Exception)
                {
                    return;
                }


                //check each style level
                for (int i = 1; i < 11; i++)
                {

                    if (styles[i] == null)
                    {
                        continue;
                    }

                    //create an arraylist of all styles

                    ArrayList AllStyles = new ArrayList();

                    //create text:index-source-styles element with appropriate text:outline-level

                    Element IndexSourceStyles = new Element("text", "index-source-styles", TEXT_NAMESPACE);
                    IndexSourceStyles.AddAttribute(new Attribute("text", "outline-level",""+i+"", TEXT_NAMESPACE));                      

                    foreach (string thisStyle in styles[i])
                    {
                        //create text:index-source-style element
                        Element IndexSourceStyleElement = new Element("text", "index-source-style", TEXT_NAMESPACE);
                        if (!AllStyles.Contains(thisStyle))
                        {
                            //add text:index-source-style element to allStyles arraylist and, as a child, to text:index-source-styles-element, if it wasn't added before
                            IndexSourceStyleElement.AddAttribute(new Attribute("text", "style-name", thisStyle, TEXT_NAMESPACE));
                            AllStyles.Add(thisStyle);
                            IndexSourceStyles.AddChild(IndexSourceStyleElement);
                        }
                    }
                        
                    IndexSourceStyles.Write(this.nextWriter);
                }
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }
        public override void WriteEndAttribute()
        {
            if (!IsIndexSourceStyleProcessed)
            {
                this.nextWriter.WriteEndAttribute();
            }
        }
        public override void WriteEndElement()
        {
            if (IsIndexSourceStyleProcessed)
            {
                this.IsIndexSourceStyleProcessed = false;
            }
            else
            {
                this.nextWriter.WriteEndElement();
            }
        }
    }
}
