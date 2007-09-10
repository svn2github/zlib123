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
using System.IO;
using CleverAge.OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.Spreadsheet
{

    public class OoxPivotCachePostProcessor : AbstractPostProcessor
    {
        private Stack pivotContext;
        private const string PXSI_NAMESPACE = "urn:cleverage:xmlns:post-processings:pivotTable";
        private const string REL_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
        private bool isPxsi;
        private bool isInPivotsNum;
        private bool isInPivotCache;
        private bool isVal;
        private bool isSheetName;
        private bool isColStart;
        private bool isColEnd;
        private bool isRowStart;
        private bool isRowEnd;
        private string[,] pivotTables; 
        //table of pivot source definitions containing {sheetName, colStart, colEnd, rowStart, rowEnd}

        private int cacheNum;


        public OoxPivotCachePostProcessor (XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.pivotContext = new Stack();
            this.isPxsi = false;
            this.isInPivotsNum = false;
            this.isInPivotCache = false;
            this.isVal = false;
            this.isSheetName = false;
            this.isColStart = false;
            this.isColEnd = false;
            this.isRowStart = false;
            this.isRowEnd = false;
            this.cacheNum = 0;
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            if (PXSI_NAMESPACE.Equals(ns) && "pivotsNum".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns));
                this.isInPivotsNum = true;
                this.isPxsi = true;
                Console.WriteLine("<pivotsNum>");
            }
            else if (PXSI_NAMESPACE.Equals(ns) && "pivotCache".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns));
                this.isInPivotCache = true;
                this.isPxsi = true;
                Console.WriteLine("<pivotCache>");
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            if (isInPivotsNum)
            {
                if (PXSI_NAMESPACE.Equals(ns) && "val".Equals(localName))
                    this.isVal = true;
            }
            else if(isInPivotCache)
            {
                if (PXSI_NAMESPACE.Equals(ns) && "sheetName".Equals(localName))
                    this.isSheetName = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "colStart".Equals(localName))
                    this.isColStart = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "colEnd".Equals(localName))
                    this.isColEnd = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "rowStart".Equals(localName))
                    this.isRowStart = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "rowEnd".Equals(localName))
                    this.isRowEnd = true;
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        public override void WriteString(string text)
        {
            if (isInPivotsNum)
            {
                if (isVal)
                {
                    this.pivotTables = new string[Convert.ToInt32(text),5];
                    Console.WriteLine("pivotsNum: " + text);
                    Console.WriteLine("pivotTables " + pivotTables.GetLength(0) + "x" + pivotTables.GetLength(1));
                    this.cacheNum = 0;
                    this.isVal = false;
                }
            }
            else if (isInPivotCache)
            {
                if (isSheetName)
                {
                    this.pivotTables[this.cacheNum,0] = text;
                    Console.WriteLine("sheetName: " + pivotTables[cacheNum, 0]);
                    this.isSheetName = false;
                }
                else if (isColStart)
                {
                    this.pivotTables[this.cacheNum,1] = text;
                    Console.WriteLine("colStart: " + pivotTables[cacheNum, 1]);
                    this.isColStart = false;
                }
                else if (isColEnd)
                {
                    this.pivotTables[this.cacheNum,2] = text;
                    Console.WriteLine("colEnd: " + pivotTables[cacheNum, 2]);
                    this.isColEnd = false;
                }
                else if (isRowStart)
                {
                    this.pivotTables[this.cacheNum,3] = text;
                    Console.WriteLine("RowStart: " + pivotTables[cacheNum, 3]);
                    this.isRowStart = false;
                }
                else if (isRowEnd)
                {
                    this.pivotTables[this.cacheNum,4] = text;
                    Console.WriteLine("rowEnd: " + pivotTables[cacheNum, 4]);
                    this.isRowEnd = false;
                }
            }
            else
            {
                this.nextWriter.WriteString(text);
            }
        }

        public override void WriteEndAttribute()
        {
            if(!isPxsi)
                this.nextWriter.WriteEndAttribute();
        }

        public override void WriteEndElement()
        {

            if (isPxsi)
            {
				Element element = (Element) pivotContext.Pop();

                if (PXSI_NAMESPACE.Equals(element.Ns) && "pivotsNum".Equals(element.Name))
                {
                    this.isInPivotsNum = false;
                    this.isPxsi = false;
                }
                else if (PXSI_NAMESPACE.Equals(element.Ns) && "pivotCache".Equals(element.Name))
                {
                    this.isInPivotCache = false;
                    this.cacheNum++;
                    this.isPxsi = false;
                }
                else
                {
                    this.nextWriter.WriteEndElement();
                }
            }
            else
            {
                this.nextWriter.WriteEndElement();
            }
        }
    }
}
