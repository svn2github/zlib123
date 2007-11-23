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
//added by sonata for mulltilevel grouping
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using OdfConverter.Transforms;
using OdfConverter.Transforms.Test;
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
            else if (text.Contains("Group-Transform"))
            {

                this.nextWriter.WriteString(EvalGroupingExpression(text));

            }
            //End
            else if (text.Contains("shade-tint"))
            {

                this.nextWriter.WriteString(EvalShadeExpression(text));

            }
            //Shadow calculation
            else if (text.Contains("shadow-offset-x") || text.Contains("shadow-offset-y"))
            {

                this.nextWriter.WriteString(EvalShadowExpression(text));
            }
            //Image Cropping Calculation Added by Sonata-15/11/2007
            else if (text.Contains("image-props"))
            {
                EvalImageCrop(text);
            }
            else
            {
                //added by clam for bug 1785583
                //Start
                if (text.StartsWith(" ") && text.Trim().Length > 0)
                {
                    try
                    {
                        this.nextWriter.WriteStartElement("text", "s", ((CleverAge.OdfConverter.OdfConverterLib.AbstractPostProcessor.Element)this.currentNode.Peek()).Ns);
                        this.nextWriter.WriteEndElement();
                        text = text.Substring(1);
                    }
                    catch (Exception ex)
                    {
                    }
                }
                //End

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
        private string EvalShadeExpression(string text)
        {
            string[] arrVal = new string[5];
            arrVal = text.Split(':');

            double dblRed = Double.Parse(arrVal[1], System.Globalization.CultureInfo.InvariantCulture);
            double dblGreen = Double.Parse(arrVal[2], System.Globalization.CultureInfo.InvariantCulture);
            double dblBlue = Double.Parse(arrVal[3], System.Globalization.CultureInfo.InvariantCulture);
            double dblShade = Double.Parse(arrVal[4], System.Globalization.CultureInfo.InvariantCulture);

            double sR;
            if (dblRed < 10)
            {
                sR = 2.4865 * dblRed;
            }
            else
            {
                sR = (Math.Pow(((dblRed + 14.025) / 269.025), 2.4) ) *8192;
            }
            double sG;
            if (dblGreen < 10)
            {
                sG = 2.4865 * dblGreen;

            }
            else {

                sG = (Math.Pow(((dblGreen + 14.025) / 269.025), 2.4) ) * 8192;
            
            }

            double sB;
            if (dblBlue < 10)
            {
                sB = 2.4865 * dblBlue;
            }
            else
            {
                sB = (Math.Pow(((dblBlue + 14.025) / 269.025), 2.4)) * 8192;
            }
           
            double NewRed;
            double sRead;
            sRead = (sR * dblShade / 100);

            if (sRead < 10)
            {
                NewRed = 0;

            }
            else if (sRead < 24)
            {
                NewRed = (12.92 * 255 * sR  / 8192);

            }
            else if (sRead <= 8192)
            {
                NewRed = ((Math.Pow(sRead / 8192, 1 / 2.4) * 1.055 )- 0.055 )* 255 ;
            }
            else
            {
                NewRed = 255;
            }
            NewRed = Math.Round(NewRed);
            
            double NewGreen;
            double sGreen;
            sGreen = (sG  * dblShade / 100);

            if (sGreen < 0)
            {
                NewGreen = 0;

            }
            else if (sGreen < 24)
            {
                NewGreen = (12.92 * 255 * dblGreen / 8192);

            }
            else if (sGreen < 8193)
            {
                NewGreen = ((Math.Pow(sGreen / 8192, 1 / 2.4) * 1.055) - 0.055) * 255;
            }
            else
            {
                NewGreen = 255;
            }
            NewGreen = Math.Round(NewGreen);

            double NewBlue;
            double sBlue;
            sBlue = (sB  * dblShade / 100);

            if (sBlue < 0)
            {
                NewBlue = 0;

            }
            else if (sBlue < 24)
            {
                NewBlue = (12.92 * 255 * dblBlue / 8192);

            }
            else if (sBlue < 8193)
            {
                NewBlue = ((Math.Pow(sBlue / 8192, 1 / 2.4) * 1.055) - 0.055) * 255; ;
            }
            else
            {
                NewBlue = 255;
            }
            NewBlue = Math.Round(NewBlue);
            
            int intRed;
            int intGreen;
            int intBlue;
            intRed = (int)NewRed;
            intGreen = (int)NewGreen;
            intBlue = (int)NewBlue;

            string hexRed;
            string hexGreen;
            string hexBlue;
            
            hexRed = String.Format("{0:x}", intRed);
            hexGreen = String.Format("{0:x}", intGreen );
            hexBlue = String.Format("{0:x}", intBlue );  
            return ('#' + hexRed.ToUpper()   + hexGreen.ToUpper ()  + hexBlue.ToUpper()  );

        }
        //Added by sonata\Vipul for multilevel grouping
        private string EvalGroupingExpression(string text)
        {
            List<OoxShape> _shapes = new List<OoxShape>();
           
            string strRet = "";
            string strShapeCordinates = "";
            string[] arrgroupShape = text.Split('$');
            string[] arrgroup = arrgroupShape[0].Split('@');

            //group cordinates
            long dblgrpX ;
            long dblgrpY ;
            long dblgrpCX ;
            long dblgrpCY;
            long dblgrpChX;
            long dblgrpChY;
            long dblgrpChCX ;
            long dblgrpChCY;
            long dblgrpRot;

            //shape cordinates
            long dblShapeX;
            long dblShapeY ;
            long dblShapeCX;
            long dblShapeCY;
            long dblShapeRot;

            OoxTransform  TopLevelgroup;
            OoxTransform Targetgroup = new OoxTransform(5495925, 3286125, 1419225, 657225, 0, 1, 1);
            OoxTransform shapeCord;
            OoxTransform InnerLevelgroup;
            OoxTransform Tempgroup;

            string[] arrShapeCordinates;
            string[] arrInnerGroup;
            string[] arrFinalRet;

            if (arrgroup.Length >= 4)
            {
                //Multilevel group

                //top level Group cordinates
                arrInnerGroup = arrgroup[1].Split(':');

                 dblgrpX = long.Parse(arrInnerGroup[0], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpY = long.Parse(arrInnerGroup[1], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpCX = long.Parse(arrInnerGroup[2], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpCY = long.Parse(arrInnerGroup[3], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpChX = long.Parse(arrInnerGroup[4], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpChY = long.Parse(arrInnerGroup[5], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpChCX = long.Parse(arrInnerGroup[6], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpChCY = long.Parse(arrInnerGroup[7], System.Globalization.CultureInfo.InvariantCulture);
                 dblgrpRot = long.Parse(arrInnerGroup[8], System.Globalization.CultureInfo.InvariantCulture);

                 TopLevelgroup = new OoxTransform(dblgrpChX, dblgrpChY, dblgrpChCX, dblgrpChCY, dblgrpX, dblgrpY, dblgrpCX, dblgrpCY, dblgrpRot, 1, 1);

                for (int intCount = 2; intCount < arrgroup.Length - 1; intCount++)
                {
                    arrInnerGroup = arrgroup[intCount].Split(':');

                    Tempgroup = TopLevelgroup;

                    //Inner level Group cordinates
                     dblgrpX = long.Parse(arrInnerGroup[0], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpY = long.Parse(arrInnerGroup[1], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpCX = long.Parse(arrInnerGroup[2], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpCY = long.Parse(arrInnerGroup[3], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpChX = long.Parse(arrInnerGroup[4], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpChY = long.Parse(arrInnerGroup[5], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpChCX = long.Parse(arrInnerGroup[6], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpChCY = long.Parse(arrInnerGroup[7], System.Globalization.CultureInfo.InvariantCulture);
                     dblgrpRot = long.Parse(arrInnerGroup[8], System.Globalization.CultureInfo.InvariantCulture);

                     InnerLevelgroup = new OoxTransform(dblgrpChX, dblgrpChY, dblgrpChCX, dblgrpChCY, dblgrpX, dblgrpY, dblgrpCX, dblgrpCY, dblgrpRot, 1, 1);

                     Targetgroup =new  OoxTransform(Tempgroup, InnerLevelgroup);

                     TopLevelgroup = Targetgroup;

                }

                arrShapeCordinates = arrgroupShape[1].Split(':'); 

                //shape cordinates
                 dblShapeX = long.Parse(arrShapeCordinates[1], System.Globalization.CultureInfo.InvariantCulture);
                 dblShapeY = long.Parse(arrShapeCordinates[2], System.Globalization.CultureInfo.InvariantCulture);
                 dblShapeCX = long.Parse(arrShapeCordinates[3], System.Globalization.CultureInfo.InvariantCulture);
                 dblShapeCY = long.Parse(arrShapeCordinates[4], System.Globalization.CultureInfo.InvariantCulture);
                 dblShapeRot = long.Parse(arrShapeCordinates[5], System.Globalization.CultureInfo.InvariantCulture);

                 _shapes.Clear();

                 shapeCord = new OoxTransform(dblShapeX, dblShapeY, dblShapeCX, dblShapeCY, dblShapeRot, 1, 1);

                 _shapes.Add(new OoxShape(new OoxTransform(Targetgroup, shapeCord)));

            }
            else if (arrgroup.Length == 3)
            {
                //  single level group
                arrInnerGroup = arrgroup[1].Split(':');

                dblgrpX = long.Parse(arrInnerGroup[0], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpY = long.Parse(arrInnerGroup[1], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpCX = long.Parse(arrInnerGroup[2], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpCY = long.Parse(arrInnerGroup[3], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpChX = long.Parse(arrInnerGroup[4], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpChY = long.Parse(arrInnerGroup[5], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpChCX = long.Parse(arrInnerGroup[6], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpChCY = long.Parse(arrInnerGroup[7], System.Globalization.CultureInfo.InvariantCulture);
                dblgrpRot = long.Parse(arrInnerGroup[8], System.Globalization.CultureInfo.InvariantCulture);

                TopLevelgroup = new OoxTransform(dblgrpChX, dblgrpChY, dblgrpChCX, dblgrpChCY, dblgrpX, dblgrpY, dblgrpCX, dblgrpCY, dblgrpRot, 1, 1);
                
                arrShapeCordinates = arrgroupShape[1].Split(':');

                //shape cordinates
                dblShapeX = long.Parse(arrShapeCordinates[1], System.Globalization.CultureInfo.InvariantCulture);
                dblShapeY = long.Parse(arrShapeCordinates[2], System.Globalization.CultureInfo.InvariantCulture);
                dblShapeCX = long.Parse(arrShapeCordinates[3], System.Globalization.CultureInfo.InvariantCulture);
                dblShapeCY = long.Parse(arrShapeCordinates[4], System.Globalization.CultureInfo.InvariantCulture);
                dblShapeRot = long.Parse(arrShapeCordinates[5], System.Globalization.CultureInfo.InvariantCulture);

                _shapes.Clear();

                shapeCord = new OoxTransform(dblShapeX, dblShapeY, dblShapeCX, dblShapeCY, dblShapeRot, 1, 1);

                _shapes.Add(new OoxShape(new OoxTransform(TopLevelgroup, shapeCord)));

            }
           
            foreach (OoxShape shape in _shapes)
            {
                strRet = shape.OoxTransform.Odf;
            }
         
            arrFinalRet = strRet.Split('@');

            if (arrFinalRet[0] == "YESROT")
            {
                if (arrgroup[0].Contains("Width"))
                {
                    if (arrFinalRet[1].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[1];
                }
                   
                }
                else if (arrgroup[0].Contains("Height"))
                {
                    if (arrFinalRet[2].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[2];
                }
                }
                else if (arrgroup[0].Contains("DrawTranform"))
                {
                    if (arrFinalRet[3].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[3];
                }
            }
            }
            else if (arrFinalRet[0] == "NOROT")
            {
                if (arrgroup[0].Contains("Width"))
                {
                    if (arrFinalRet[1].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[1];
                }
                }
                else if (arrgroup[0].Contains("Height"))
                {
                    if (arrFinalRet[2].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[2];
                }
                }
                else if (arrgroup[0].Contains("SVGX"))
                {
                    if (arrFinalRet[3].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[3];
                }
                }
                else if (arrgroup[0].Contains("SVGY"))
                {
                    if (arrFinalRet[4].Contains("NaN"))
                    {
                        strShapeCordinates = "0cm";
                    }
                    else
                    {
                    strShapeCordinates = arrFinalRet[4];
                }
            }
            }


            return strShapeCordinates;

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

        //Image Cropping Added by Sonata-15/11/2007
        private void EvalImageCrop(string text)
        {
            string[] arrVal = new string[6];
            arrVal = text.Split(':');
            string source = arrVal[1].ToString();
            int left = int.Parse(arrVal[2].ToString(),System.Globalization.CultureInfo.InvariantCulture);
            int right = int.Parse(arrVal[3].ToString(),System.Globalization.CultureInfo.InvariantCulture);
            int top = int.Parse(arrVal[4].ToString(),System.Globalization.CultureInfo.InvariantCulture);
            int bottom = int.Parse(arrVal[5].ToString(),System.Globalization.CultureInfo.InvariantCulture);


            string tempFileName = AbstractConverter.inputTempFileName.ToString();
            ZipResolver zipResolverObj = new ZipResolver(tempFileName);
            OdfZipUtils.ZipArchiveWriter zipobj = new OdfZipUtils.ZipArchiveWriter(zipResolverObj);
            string widht_height_res = zipobj.ImageCopyBinary(source);
            zipResolverObj.Dispose();
            zipobj.Close();


            string[] arrValues = new string[3];
            arrValues = widht_height_res.Split(':');
            double width = double.Parse(arrValues[0].ToString(),System.Globalization.CultureInfo.InvariantCulture);
            double height = double.Parse(arrValues[1].ToString(),System.Globalization.CultureInfo.InvariantCulture);
            double res = double.Parse(arrValues[2].ToString(),System.Globalization.CultureInfo.InvariantCulture);


            double cx = width * 2.54 / res;
            double cy = height * 2.54 / res;

            double odpLeft = left * cx / 100000;
            double odpRight = right * cx / 100000;
            double odpTop = top * cy / 100000;
            double odpBottom = bottom * cy / 100000;

            string result = string.Concat("rect(", string.Format("{0:0.##}", odpTop) + "cm" + " " + string.Format("{0:0.##}", odpRight) + "cm" + " " + string.Format("{0:0.##}", odpBottom) + "cm" + " " + string.Format("{0:0.##}", odpLeft) + "cm", ")");
            this.nextWriter.WriteString(result);

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
