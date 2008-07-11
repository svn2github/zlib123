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
/*
Modification Log
LogNo. |Date       |ModifiedBy   |BugNo.   |Modification                                                      |
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RefNo-1 22-Jan-2008 Sandeep S     1832335   to avoid New line inserted in note content after roundtrip conversions                 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
*/

using System;
using System.Xml;
using System.Collections;
using CleverAge.OdfConverter.OdfConverterLib;
using CleverAge.OdfConverter.OdfZipUtils;


namespace CleverAge.OdfConverter.Spreadsheet
{
    /// <summary>
    /// Postprocessor which allows to move sharedstrings into cells.
    /// </summary>
    public class OdfConditionalPostProcessor : AbstractPostProcessor
    {
       
        private const string PXSI_NAMESPACE = "urn:cleverage:xmlns:post-processings:shared-strings";  
        private Element sharedStringElement;

        private string ElementName = "";
        private Hashtable stringSqref;
        private string StyleName;
        private string StringValue;

        public OdfConditionalPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {          
            
            stringSqref = new Hashtable();

        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            if ("ConditionalCell".Equals(localName))
            {
                ElementName = "ConditionalCell";
            }
            else if ("conditionalFormatting".Equals(localName))
            {
                ElementName = "conditionalFormatting";
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
            

        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            if ("StyleName".Equals(localName))
            {
                ElementName = "StyleName";                 
            }
            else if (ElementName.Equals("ConditionalCell"))
            {
               
            }            
            
            else
            {
                if ("sqref".Equals(localName) && ElementName.Equals("conditionalFormatting"))
                {
                    ElementName = "sqref";
                }
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
            

        }

        public override void WriteString(string text)
        {
            if (ElementName.Equals("ConditionalCell"))
            {
                StringValue = text;
            }
            else if (ElementName.Equals("StyleName"))
            {
                StyleName = text;
            }
            else if (ElementName.Equals("sqref"))
            {
                this.nextWriter.WriteString(stringSqref[text].ToString());              
                ElementName = "";
            }
            else
            {
                this.nextWriter.WriteString(text);
            }

        }

        public override void WriteEndAttribute()
        {
            if (ElementName.Equals("StyleName"))
            {
                ElementName = "ConditionalCell";
            }
            else if (ElementName.Equals("ConditionalCell"))
            {

            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }
        }

        public override void WriteEndElement()
        {
            if (ElementName.Equals("ConditionalCell"))
            {
                ElementName = "";
                if (!stringSqref.ContainsKey(StyleName))
                {
                    stringSqref.Add(StyleName, StringValue);
                }
                else
                {
                    stringSqref[StyleName] = stringSqref[StyleName] + " " + StringValue;
                }
            }
            else
            {
                this.nextWriter.WriteEndElement();
            }
            
        }
    }
}
