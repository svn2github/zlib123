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
using CleverAge.OdfConverter.OdfZipUtils;


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
        private bool isPxsiNumber;
        private bool v;
        private bool sst;
        private Hashtable stringCellValue;

        public OdfSharedStringsPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.sharedStringContext = new Stack();
            this.isInSharedString = false;
            this.isString = false;            
            this.isPxsi = false;
            this.isPxsiNumber = false;
            this.v = false;
            stringCellValue = new Hashtable();

        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {

            if ("si".Equals(localName))
            {
                this.isPxsi = true;
            }
            
            else if ("v".Equals(localName))
            {
                this.v = true;
            }
            else if ("sst".Equals(localName))
            {
                this.sst = true;
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }


        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            if (this.isPxsi)
            {
                if ("number".Equals(localName))
                {
                    this.isPxsiNumber = true;
                }
            }
            else if (this.v)
            {
            }
            else if (this.sst)
            {
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }

        }

        public override void WriteString(string text)
        {
            if (this.isPxsiNumber)
            {
                this.stringNumber = text;
            }            
            else if (this.isPxsi)
            {
                if (!stringCellValue.ContainsKey(this.stringNumber))
                {
                    stringCellValue.Add(this.stringNumber, text);                    
                }
            }
            else if (this.v)
            {
                //int index = Convert.ToInt32(text);

                if (stringCellValue.ContainsKey(text))
                {                    
                    this.nextWriter.WriteString(stringCellValue[text].ToString());
                }
                
            }
            else if (this.sst)
            {
            }
               
            /*
              Bug No.         :1805599
              Code Modified By:Vijayeta
              Date            :6th Nov '07
              Description     :If New Line(\n) is present in the text content, Insert empty <text:p> nodes
           */

            else if (text.Contains("SonataAnnotation"))
            {

                string[] content = text.Split('|');
                string style = content[content.Length - 1];
                string textContent = content[1];
                if (textContent.Contains("\n"))
                {
                    char[] a = new char[] { '\n' };
                    string[] p = textContent.Split(a);
                    for (int i = 0; i < p.Length; i++)
                    {
                        if (p[i] == "")
                        {
                            WriteStartElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                            WriteStartElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                            WriteEndElement();
                            WriteEndElement();
                        }
                        if (p[i] != "")
                        {
                            WriteStartElement("text", "p", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                            WriteStartElement("text", "span", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                            nextWriter.WriteStartAttribute("text", "style-name", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                            nextWriter.WriteValue(style);
                            nextWriter.WriteEndAttribute();
                            this.nextWriter.WriteString(p[i]);
                            WriteEndElement();
                            WriteEndElement();

                        }
                    }

                }
            }
            //End of modification for the bug 1805599

             // Image Cropping   Added by Sonata
            else if (text.Contains("image-props"))
            {
                string[] arrVal = new string[6];
                arrVal = text.Split(':');
                string source = arrVal[1].ToString();
                int left = int.Parse(arrVal[2].ToString(),System.Globalization.CultureInfo.InvariantCulture);
                int right = int.Parse(arrVal[3].ToString(),System.Globalization.CultureInfo.InvariantCulture);
                int top = int.Parse(arrVal[4].ToString(),System.Globalization.CultureInfo.InvariantCulture);
                int bottom = int.Parse(arrVal[5].ToString(),System.Globalization.CultureInfo.InvariantCulture);

                string tempFileName = AbstractConverter.inputTempFileName.ToString();
                ZipResolver resolverObj = new ZipResolver(tempFileName);
                ZipArchiveWriter zipObj = new ZipArchiveWriter(resolverObj);
                string imgaeValues = zipObj.ImageCopyBinary(source);
                zipObj.Close();
                resolverObj.Dispose();

                string[] arrValues = new string[3];
                arrValues = imgaeValues.Split(':');
                double width = double.Parse(arrValues[0].ToString(),System.Globalization.CultureInfo.InvariantCulture);
                double height = double.Parse(arrValues[1].ToString(),System.Globalization.CultureInfo.InvariantCulture);
                double res = double.Parse(arrValues[2].ToString(),System.Globalization.CultureInfo.InvariantCulture);


                double cx = width * 2.54 / res;
                double cy = height * 2.54 / res;

                double odpLeft = (left * cx / 100000)/2.54;
                double odpRight = (right * cx / 100000)/2.54;
                double odpTop = (top * cy / 100000)/2.54;
                double odpBottom = (bottom * cy / 100000)/2.54;

                string result = string.Concat("rect(", string.Format(System.Globalization.CultureInfo.InvariantCulture,"{0:0.##}", odpTop) + "in" + " " + string.Format(System.Globalization.CultureInfo.InvariantCulture,"{0:0.##}", odpRight) + "in" + " " + string.Format(System.Globalization.CultureInfo.InvariantCulture,"{0:0.##}", odpBottom) + "in" + " " + string.Format(System.Globalization.CultureInfo.InvariantCulture,"{0:0.##}", odpLeft) + "in", ")");
                this.nextWriter.WriteString(result);


              
            }
         
            else
            {

                this.nextWriter.WriteString(text);
            }


        }

        public override void WriteEndAttribute()
        {
            if (this.isPxsiNumber)
            {
                this.isPxsiNumber = false;
            }
            else if (this.isPxsi)
            {
                
            }
           
            else if (this.v)
            {
            }
            else if (this.sst)
            {
            }
            else
            {
                this.nextWriter.WriteEndAttribute();
            }
           


        }

        public override void WriteEndElement()
        {
            if (this.isPxsi)
            {
                this.isPxsi = false;
            }
            else if (this.v)
            {
                this.v = false;
            }
            else if (this.sst)
            {
                this.sst = false;
            }

            else
            {
                this.nextWriter.WriteEndElement();
            }



        }
    }
}
