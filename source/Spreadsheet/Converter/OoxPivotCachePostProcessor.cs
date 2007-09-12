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
        private const string EXCEL_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private bool isPxsi;

        private bool isInPivotTable;        
        private bool isSheetNum;
        private string cacheSheetNum;
        private bool isColStart;
        private string cacheColStart;
        private bool isColEnd;
        private string cacheColEnd;
        private bool isRowStart;
        private string cacheRowStart;
        private bool isRowEnd;
        private string cacheRowEnd;
        private string[,] pivotTable;
        private Hashtable[] fieldItems;

        private bool isInPivotCell;
        private Hashtable pivotCells;
        private bool isPivotCellSheet;
        private string pivotCellSheet;
        private bool isPivotCellCol;
        private string pivotCellCol;
        private bool isPivotCellRow;
        private string pivotCellRow;
        private string pivotCellVal;

        
        public OoxPivotCachePostProcessor (XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.pivotContext = new Stack();
            this.isPxsi = false;

            this.isInPivotTable = false;
            this.isSheetNum = false;
            this.cacheSheetNum = "";
            this.isColStart = false;
            this.cacheColStart = "";
            this.isColEnd = false;
            this.cacheColEnd = "";
            this.isRowStart = false;
            this.cacheRowStart = "";
            this.isRowEnd = false;
            this.cacheRowEnd = "";

            this.isInPivotCell = false;
            this.isPivotCellSheet = false;
            this.pivotCellSheet = "";
            this.isPivotCellCol = false;
            this.pivotCellCol = "";
            this.isPivotCellRow = false;
            this.pivotCellRow = "";
            this.pivotCellVal = "";
            this.pivotCells = new Hashtable();
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {

            if (PXSI_NAMESPACE.Equals(ns) && "pivotCell".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns));
                this.isInPivotCell = true;
                this.isPxsi = true;
                //Console.WriteLine("<PivotCell>");
            }
            else if (PXSI_NAMESPACE.Equals(ns) && "pivotTable".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns));
                this.isInPivotTable = true;
                this.isPxsi = true;
                //Console.WriteLine("<PivotTable>");
            }
            else
            {
                this.nextWriter.WriteStartElement(prefix, localName, ns);
            }
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            if (isInPivotTable)
            {
                if (PXSI_NAMESPACE.Equals(ns) && "sheetNum".Equals(localName))
                    this.isSheetNum = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "colStart".Equals(localName))
                    this.isColStart = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "colEnd".Equals(localName))
                    this.isColEnd = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "rowStart".Equals(localName))
                    this.isRowStart = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "rowEnd".Equals(localName))
                    this.isRowEnd = true;
            }
            else if (isInPivotCell)
            {
                if (PXSI_NAMESPACE.Equals(ns) && "col".Equals(localName))
                    this.isPivotCellCol = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "row".Equals(localName))
                    this.isPivotCellRow = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "sheetNum".Equals(localName))
                    this.isPivotCellSheet = true;
            }
            else
            {
                this.nextWriter.WriteStartAttribute(prefix, localName, ns);
            }
        }

        public override void WriteString(string text)
        {
            if (isInPivotTable)
            {
                if (isSheetNum)
                {
                    this.cacheSheetNum = text;
                    //Console.WriteLine("cacheSheetNum: " + this.cacheSheetNum);
                    this.isSheetNum = false;
                }
                else if (isColStart)
                {
                    this.cacheColStart = text;
                    //Console.WriteLine("cacheColStart: " + this.cacheColStart);
                    this.isColStart = false;
                }
                else if (isColEnd)
                {
                    this.cacheColEnd = text;
                    //Console.WriteLine("cacheColEnd: " + this.cacheColEnd);
                    this.isColEnd = false;
                }
                else if (isRowStart)
                {
                    this.cacheRowStart = text;
                    //Console.WriteLine("cacheRowStart: " + this.cacheRowStart);
                    this.isRowStart = false;
                }
                else if (isRowEnd)
                {
                    this.cacheRowEnd = text;
                    //Console.WriteLine("cacheRowStart: " + this.cacheRowEnd);
                    this.isRowEnd = false;
                }
            }
            else if (isInPivotCell)
            {
                if (isPivotCellCol)
                {
                    this.pivotCellCol = text;
                    //Console.WriteLine("PivotCellCol: " + this.PivotCellCol);
                    this.isPivotCellCol = false;
                }
                else if (isPivotCellRow)
                {
                    this.pivotCellRow = text;
                    //Console.WriteLine("PivotCellRow: " + this.PivotCellRow);
                    this.isPivotCellRow = false;
                }
                else if (isPivotCellSheet)
                {
                    this.pivotCellSheet = text;
                    ////Console.WriteLine("PivotCellSheet: " + this.PivotCellSheet);
                    this.isPivotCellSheet = false;
                }
                else
                {
                    this.pivotCellVal = text;
                    //Console.WriteLine("PivotCellRow: " + this.PivotCellVal);
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

                if (PXSI_NAMESPACE.Equals(element.Ns) && "pivotTable".Equals(element.Name))
                {
                    Console.WriteLine("<SHEET " + cacheSheetNum + ">");

                    this.pivotTable = new string[Convert.ToInt32(cacheRowEnd) - Convert.ToInt32(cacheRowStart) + 1, Convert.ToInt32(cacheColEnd) - Convert.ToInt32(cacheColStart) + 1];
                    CreatePivotTable(Convert.ToInt32(cacheSheetNum),Convert.ToInt32(cacheRowStart),Convert.ToInt32(cacheRowEnd),Convert.ToInt32(cacheColStart),Convert.ToInt32(cacheColEnd));

                    this.isInPivotTable = false;
                    this.cacheSheetNum = "";
                    this.cacheColStart = "";
                    this.cacheColEnd = "";
                    this.cacheRowStart = "";
                    this.cacheRowEnd = "";
                    this.isPxsi = false;
                }

                else if (PXSI_NAMESPACE.Equals(element.Ns) && "pivotCell".Equals(element.Name))
                {
                    pivotCells.Add(pivotCellSheet + "#" + pivotCellCol + ":" + pivotCellRow, pivotCellVal);

                    this.isInPivotCell = false;
                    this.pivotCellCol = "";
                    this.pivotCellRow = "";
                    this.pivotCellVal = "";
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

        private void CreatePivotTable(int sheetNum, int rowStart, int rowEnd, int colStart, int colEnd)
        {

            int rows = rowEnd - rowStart + 1;
            int cols = colEnd - colStart + 1;
            int[] index = new int[cols];

            this.fieldItems = new Hashtable[cols];

            for (int field = 0; field < colEnd - colStart + 1; field++)
                fieldItems[field] = new Hashtable();
            
            for (int row = 0; row < rows; row++)
                for (int col = 0; col < cols; col++)
                {
                    int keyCol = Convert.ToInt32(colStart) + col;
                    int keyRow = Convert.ToInt32(rowStart) + row;

                    string key = sheetNum + "#" + keyCol.ToString() + ":" + keyRow.ToString();

                    if (pivotCells.ContainsKey(key))
                    {                        
                        this.pivotTable[row, col] = (string)pivotCells[key];

                        if (!fieldItems[col].ContainsKey((string)pivotCells[key]))
                        {
                            fieldItems[col].Add((string)pivotCells[key], index[col]);
                            index[col]++;
                        }
                    }
                    else
                    {
                        this.pivotTable[row, col] = "";
                        if (!fieldItems[col].ContainsKey(""))
                        {
                            fieldItems[col].Add("", index[col]);
                            index[col]++;
                        }
                    }
                }

            //Console.WriteLine(indeces[1].Count);
                foreach (string key in this.fieldItems[0].Keys)
                {
                    Console.WriteLine("-" + fieldItems[0][key] + ":" + key);
                }

        }
    }
}
