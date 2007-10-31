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
using System.IO;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for characters post processings
    public class OdfCharactersPostProcessor : AbstractPostProcessor
    {

        private const string TEXT_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private const string PCHAR_NAMESPACE = "urn:cleverage:xmlns:post-processings:characters";

        private Stack currentNode;
        private Stack store;

        public OdfCharactersPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.currentNode = new Stack();
            this.store = new Stack();
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            Element e = null;
            if (TEXT_NAMESPACE.Equals(ns) && "span".Equals(localName))
            {
                e = new Span(prefix, localName, ns);
            }
            else
            {
                e = new Element(prefix, localName, ns);
            }

            this.currentNode.Push(e);

            if (InSpan())
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

            if (IsSpan())
            {
                WriteStoredSpan();
            }
            if (InSpan())
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

            if (InSpan())
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
            if (InSpan())
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
            //added by clam for bug 1785583
            //Start
            //if (text.StartsWith(" "))
            //{
            //    try
            //    {
            //        this.nextWriter.WriteStartElement("text", "s", ((CleverAge.OdfConverter.OdfConverterLib.AbstractPostProcessor.Element)this.currentNode.Peek()).Ns);
            //        this.nextWriter.WriteEndElement();
            //        text = text.Substring(1);
            //    }
            //    catch (Exception ex)
            //    {                   
            //    }
            //}
            //End
          
           //for getting system date  (text ==":::current:::")
           if (InSpan() && (!(text ==":::current:::")))
            {
                StoreString(text);
            }
            else if (text.Contains("svg-x") || text.Contains("svg-y"))
            {

                this.nextWriter.WriteString(EvalExpression(text));

            }
           
           else if (text ==":::current:::")
            {

                this.nextWriter.WriteString(EvalgetCurrentSysDate(text));

            }
            // added by vipul for Shape Rotation
            //Start
            else if (text.Contains("draw-transform"))
            {

                this.nextWriter.WriteString(EvalRotationExpression(text));

            }
            //End
            //Shadow calculation
            else if (text.Contains("shadow-offset-x") || text.Contains("shadow-offset-y"))
            {

                this.nextWriter.WriteString(EvalShadowExpression(text));
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }
        /*
         * General methods */
        private string EvalgetCurrentSysDate(string text)
        {
            string str = "";
            str = DateTime.Now.ToShortDateString();
            return str;

        }
        private string EvalExpression(string text)
        {
            string[] arrVal = new string[4];
            arrVal = text.Split(':');
            Double x = 0;
            if (arrVal.Length == 5)
            {
                double arrValue1 = Double.Parse(arrVal[1], System.Globalization.CultureInfo.InvariantCulture);
                double arrValue2 = Double.Parse(arrVal[2], System.Globalization.CultureInfo.InvariantCulture);
                double arrValue3 = Double.Parse(arrVal[3], System.Globalization.CultureInfo.InvariantCulture);
                double arrValue4 = Double.Parse(arrVal[4], System.Globalization.CultureInfo.InvariantCulture);

                if (arrVal[0].Contains("svg-x1"))
                {
                    x = Math.Round(((arrValue1 -
                        Math.Cos(arrValue4) * arrValue2 +
                        Math.Sin(arrValue4) * arrValue3) / 360000), 2);
                }
                else if (arrVal[0].Contains("svg-y1"))
                {
                    x = Math.Round(((arrValue1 -
                        Math.Sin(arrValue4) * arrValue2 -
                        Math.Cos(arrValue4) * arrValue3) / 360000), 2);
                }
                else if (arrVal[0].Contains("svg-x2"))
                {
                    x = Math.Round(((arrValue1 +
                        Math.Cos(arrValue4) * arrValue2 -
                        Math.Sin(arrValue4) * arrValue3) / 360000), 2);
                }
                else if (arrVal[0].Contains("svg-y2"))
                {
                    x = Math.Round(((arrValue1 +
                        Math.Sin(arrValue4) * arrValue2 +
                        Math.Cos(arrValue4) * arrValue3) / 360000), 2);
                }
                
            }

            return x.ToString().Replace(',','.') + "cm";
        }
        // added by vipul for Shape Rotation
        //Start
        private string EvalRotationExpression(string text)
        {
            string[] arrVal = new string[7];
            string strRotate;
            string strTranslate;

            arrVal = text.Split(':');
            
            Double dblRadius = 0.0;
            Double dblXf = 0.0;
            Double dblYf = 0.0;
            Double dblalpha = 0.0;
            Double dblbeta= 0.0;
            Double dblgammaDegree = 0.0;
            Double dblgammaR = 0.0;
            Double dblRotAngle = 0.0;
            Double dblX2 = 0.0;
            Double dblY2 = 0.0;
            Double centreX = 0.0;
            Double centreY = 0.0;
            
            if (arrVal.Length == 8)
            {
                double arrValueX = Double.Parse(arrVal[1], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueY = Double.Parse(arrVal[2], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueCx = Double.Parse(arrVal[3], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueCy = Double.Parse(arrVal[4], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueFlipH = Double.Parse(arrVal[5], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueFlipV = Double.Parse(arrVal[6], System.Globalization.CultureInfo.InvariantCulture);
                double arrValueRot = Double.Parse(arrVal[7], System.Globalization.CultureInfo.InvariantCulture);

               

                if (arrVal[0].Contains("draw-transform"))
                {

                    centreX = arrValueX + (arrValueCx /2);
                    centreY = arrValueY + (arrValueCy / 2);

                    if (arrValueFlipH == 1.0)
                    {
                        dblXf = arrValueX + ((centreX - arrValueX) * 2);
                    }
                    else
                    {
                        dblXf = arrValueX; 
                    }

                    if (arrValueFlipV == 1.0)
                    {
                        dblYf = arrValueY + ((centreY - arrValueY) * 2);
                    }
                    else
                    {
                        dblYf = arrValueY;
                    }
                    dblRadius = Math.Sqrt((arrValueCx * arrValueCx) + (arrValueCy * arrValueCy)) / 2.0;
                    
                    if ( (arrValueFlipH ==0.0 && arrValueFlipV == 1.0 ) || ( arrValueFlipH ==0.0 && arrValueFlipV ==1.0 ) )
                    {
                        dblalpha = 360.00 - ( (arrValueRot / 60000.00 ) % 360 ) ;
                    }
                    else
                    {
                         dblalpha =(arrValueRot / 60000.00 ) % 360 ;
                    }
                    if (dblalpha > 180.00)
                    {
                        dblRotAngle = (360.00 - dblalpha) / 180.00 * Math.PI;
                    }
                    else
                    {
                        dblRotAngle = (-1.00 * dblalpha) / 180.00 * Math.PI;
                    }
                    if (Math.Abs(centreY - dblYf) > 0)
                    {
                    dblbeta = Math.Atan(Math.Abs(centreX - dblXf) / Math.Abs(centreY - dblYf)) * (180.00 / Math.PI);
                    }

                    if (Math.Abs(dblbeta - dblalpha) > 0)
                    {
                        dblgammaDegree = ((dblbeta - dblalpha) % ((int)((dblbeta - dblalpha) / Math.Abs(dblbeta - dblalpha)) * 360)) + 90.00;
                    }
                    else
                    {
                        dblgammaDegree = 90.00;
                    }
                                       
                    dblgammaR = (360.00 - dblgammaDegree) / 180.00 * Math.PI;
                    dblX2 = Math.Round((centreX + (dblRadius * Math.Cos(dblgammaR))) / 360036.00, 3);
                    dblY2 = Math.Round((centreY + (dblRadius * Math.Sin(dblgammaR))) / 360036.00, 3);
             
               } 

            }
                strRotate="rotate (" + dblRotAngle.ToString().Replace(',', '.') +")";
                          
                strTranslate = "translate (" + dblX2.ToString().Replace(',', '.') + "cm " + dblY2.ToString().Replace(',', '.') + "cm)";

                 return strRotate+" "+strTranslate;
        }
        //End 

        //Resolve relative path to absolute path
        private string HyperLinkPath(string text)
        {
            string[] arrVal = new string[2];
            arrVal = text.Split(':');
            string source = arrVal[1].ToString();
            string address = "";

            if (arrVal.Length == 2)
            {
                string returnInputFilePath = "";

                // for Commandline tool
                for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
                {
                    if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/I")
                        returnInputFilePath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
                }

                //for addins
                if (returnInputFilePath == "")
                {
                    returnInputFilePath = Environment.CurrentDirectory;
                }
                
                string linkPathLocation = Path.GetFullPath(Path.Combine(returnInputFilePath, source)).Replace("\\", "/").Replace(" ","%20");
                address = "/" + linkPathLocation;
               
            }
            return address.ToString();
           
        }
        //End

        // added for Shadow calculation
        private string EvalShadowExpression(string text)
        {
            string[] arrVal = new string[2];
            arrVal = text.Split(':');
            Double x = 0;
            if (arrVal.Length == 3)
            {
                double arrDist = Double.Parse(arrVal[1], System.Globalization.CultureInfo.InvariantCulture);
                double arrDir = Double.Parse(arrVal[2], System.Globalization.CultureInfo.InvariantCulture);

                arrDir = (arrDir / 60000) * (Math.PI / 180.0);
                if (arrVal[0].Contains("shadow-offset-x"))
                {
                    x = Math.Sin((arrDir)) * (arrDist / 360000);
                }
                else if (arrVal[0].Contains("shadow-offset-y"))
                {
                    x = Math.Cos((arrDir)) * (arrDist / 360000);
                }

            }

            return x.ToString() + "cm";

        }
        public void WriteStoredSpan()
        {
            Element e = (Element)this.store.Peek();

            if (e is Span)
            {
                Span span = (Span)e;

                if (span.HasSymbol)
                {
                    string code = span.GetChild("symbol", PCHAR_NAMESPACE).GetAttributeValue("code", PCHAR_NAMESPACE);
                    span.Replace(span.GetChild("symbol", PCHAR_NAMESPACE), char.ConvertFromUtf32(int.Parse(code, System.Globalization.NumberStyles.HexNumber)));
                    // do not write the span if embedded in another span (it will be written afterwards)
                    if (!IsEmbeddedSpan())
                    {
                        span.Write(nextWriter);
                    }
                }
                else
                {
                    if (this.store.Count < 2)
                    {
                        span.Write(nextWriter);
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
            Element element = (Element)this.currentNode.Peek();

            if (this.store.Count > 0)
            {
                Element parent = (Element)this.store.Peek();
                parent.AddChild(element);
            }

            this.store.Push(element);
        }


        private void EndStoreElement()
        {
            Element e = (Element)this.store.Pop();
        }


        private void StartStoreAttribute()
        {
            Element parent = (Element)store.Peek();
            Attribute attr = (Attribute)this.currentNode.Peek();
            parent.AddAttribute(attr);
            this.store.Push(attr);
        }


        private void StoreString(string text)
        {
            Node node = (Node)this.currentNode.Peek();

            if (node is Element)
            {
                /* This condition is to check tab carector in a string and 
                replace with text:tab node */
                if (text.Contains("\t"))
                {
                    char[] a = new char[] { '\t' };
                    string[] p = text.Split(a);

                    for (int i = 0; i < p.Length; i++)
                    {
                        Element element = (Element)this.store.Peek();
                        element.AddChild(p[i].ToString());
                        if (i < p.Length-1)
                        {                            
                            WriteStartElement("text","tab" ,element.Ns);
                            WriteEndElement();                            
                        }
                    }
                }                
                else
                {
                    Element element = (Element)this.store.Peek();

                    element.AddChild(text);
                }

            }
            else
            {
                Attribute attr = (Attribute)store.Peek();
                // This condition is to check for hyperlink relative path 
                if (text.Contains("hyperlink-path"))
                {
                    string hPath = HyperLinkPath(text);
                    attr.Value += hPath;
                }
                else
                {
                attr.Value += text;
            }
                
            }
        }       

        private void EndStoreAttribute()
        {
            this.store.Pop();
        }

        private bool IsSpan()
        {
            Node node = (Node)this.currentNode.Peek();
            if (node is Element)
            {
                Element element = (Element)node;
                if ("span".Equals(element.Name) && TEXT_NAMESPACE.Equals(element.Ns))
                {
                    return true;
                }
            }
            return false;
        }


        private bool InSpan()
        {
            return IsSpan() || this.store.Count > 0;
        }


        private bool IsEmbeddedSpan()
        {
            Node node = (Node)this.currentNode.ToArray().GetValue(1);
            if (node is Element)
            {
                Element element = (Element)node;
                if ("span".Equals(element.Name) && TEXT_NAMESPACE.Equals(element.Ns))
                {
                    return true;
                }
            }
            return false;
        }


        protected class Span : Element
        {
            public Span(Element e)
                :
                base(e.Prefix, e.Name, e.Ns) { }

            public Span(string prefix, string localName, string ns)
                :
                base( prefix, localName, ns) { }

            public bool HasSymbol
            {
                get
                {
                    return HasSymbolNode(this);
                }
            }

            private bool HasSymbolNode(Element e)
            {
                if (e.GetChild("symbol", PCHAR_NAMESPACE) != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }


    }
}
