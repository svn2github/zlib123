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

        //<pxsi:pivotTable> variables
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
        private Hashtable[,] fieldItems;
        //fieldItems[i,j] (dimesion i equals number of pivotFields, dimesion j equals 2) 
        //fieldItems variable for each pivotField contains hashtable with unique values of this pivotField and a sequential number
        //at fieldItems[i,0] pivotField[i] values are the key and sequential number are the key values
        //at fieldItems[i,1] sequential number is the key and pivotField[i] values are the key values  

        //<pxsi:pivotCell> variables
        private bool isInPivotCell;
        private Hashtable pivotCells;
        private bool isPivotCellSheet;
        private string pivotCellSheet;
        private bool isPivotCellCol;
        private string pivotCellCol;
        private bool isPivotCellRow;
        private string pivotCellRow;
        private string pivotCellVal;

        //<pxsi:sharedItems> variables
        private bool isInSharedItems;
        private bool isFieldType;
        private string fieldType; //{"axis" or "data"}
        private bool isFieldNum;
        private string fieldNum;

        //<pxsi:cacheRecords> variables
        private bool isInCacheRecords;

        public OoxPivotCachePostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
        {
            this.pivotContext = new Stack();
            this.isPxsi = false;

            //<pxsi:pivotTable> variables
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

            //<pxsi:pivotCell> variables
            this.isInPivotCell = false;
            this.isPivotCellSheet = false;
            this.pivotCellSheet = "";
            this.isPivotCellCol = false;
            this.pivotCellCol = "";
            this.isPivotCellRow = false;
            this.pivotCellRow = "";
            this.pivotCellVal = "";
            this.pivotCells = new Hashtable();

            //<pxsi:sharedItems> variables
            this.isInSharedItems = false; ;
            this.isFieldType = false; ;
            this.fieldType = "";
            this.isFieldNum = false;
            this.fieldNum = "";

            //<pxsi:cacheRecords> variables
            this.isInCacheRecords = false;

            Console.WriteLine("PROCESS");

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
            else if (PXSI_NAMESPACE.Equals(ns) && "sharedItems".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns));
                this.isInSharedItems = true;
                this.isPxsi = true;
            }
            else if (PXSI_NAMESPACE.Equals(ns) && "cacheRecords".Equals(localName))
            {
                this.pivotContext = new Stack();
                this.pivotContext.Push(new Element(prefix, localName, ns)); 
                this.isInCacheRecords = true;
                this.isPxsi = true;
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
            else if (isInSharedItems)
            {
                if (PXSI_NAMESPACE.Equals(ns) && "fieldType".Equals(localName))
                    this.isFieldType = true;
                else if (PXSI_NAMESPACE.Equals(ns) && "fieldNum".Equals(localName))
                    this.isFieldNum = true;
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
            else if (isInSharedItems)
            {
                if (isFieldType)
                {
                    this.fieldType = text;
                    this.isFieldType = false;
                }
                else if (isFieldNum)
                {
                    this.fieldNum = text;
                    this.isFieldNum = false;
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
                    
                    //other <pivotTable> variables are being used in other postprocessor commands
                    //so they are not erased here

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
                else if (PXSI_NAMESPACE.Equals(element.Ns) && "sharedItems".Equals(element.Name))
                {
                    this.nextWriter.WriteStartElement("sharedItems", EXCEL_NAMESPACE);

                    int field = Convert.ToInt32(fieldNum);
                   
                    //insert "count" attribute
                    this.nextWriter.WriteStartAttribute("count");
                    this.nextWriter.WriteString((string) this.fieldItems[field,0].Count.ToString());
                    this.nextWriter.WriteEndAttribute();
                    
                    //SORTOWANIE
                    //string[] uniqueValues = new string[this.fieldItems[field, 0].Count];
                    //for (int i = 0; i < this.fieldItems[field, 1].Count; i++)
                    //    uniqueValues[i] = this.fieldItems[field, 1][i].ToString();
                    //Array.Sort(uniqueValues);
                    //Console.WriteLine("FIELD");
                    //for (int i = 0; i < uniqueValues.Length; i++)
                    //    Console.WriteLine(uniqueValues[i]);
                    

                    //write item elements
                    if ("axis".Equals(this.fieldType))
                        for (int i = 0; i < this.fieldItems[field, 1].Count; i++ )
                            //foreach (string key in this.fieldItems[Convert.ToInt32(fieldNum), 0].Keys)
                            {
                                try
                                {
                                    //if field value is a number
                                    Convert.ToDouble(this.fieldItems[field, 1][i]);
                                    this.nextWriter.WriteStartElement("n", EXCEL_NAMESPACE);
                                }
                                catch
                                {
                                    //if field value is a string
                                    this.nextWriter.WriteStartElement("s", EXCEL_NAMESPACE);
                                }
                             
                                this.nextWriter.WriteStartAttribute("v");
                                this.nextWriter.WriteString(this.fieldItems[field,1][i].ToString());
                                this.nextWriter.WriteEndAttribute();
                                this.nextWriter.WriteEndElement();
                            }

                    this.nextWriter.WriteEndElement();
                    this.isInSharedItems = false;
                    this.fieldNum = "";
                    this.fieldType = "";
                    this.isPxsi = false;
                }
                else if (isInCacheRecords)
                {
                    string value;

                    int rows = Convert.ToInt32(cacheRowEnd) - Convert.ToInt32(cacheRowStart) + 1;
                    int cols = Convert.ToInt32(cacheColEnd) - Convert.ToInt32(cacheColStart) + 1;

                    for (int row = 0; row < rows; row++)
                    {
                        this.nextWriter.WriteStartElement("r",EXCEL_NAMESPACE);
                        for (int col = 0; col < cols; col++)
                        {
                            try
                            {
                                //if this is number value
                                Convert.ToDouble(pivotTable[row, col].Replace('.',','));
                                this.nextWriter.WriteStartElement("n",EXCEL_NAMESPACE);
                                this.nextWriter.WriteStartAttribute("v");
                                this.nextWriter.WriteString(pivotTable[row, col]);
                                this.nextWriter.WriteEndAttribute();
                                this.nextWriter.WriteEndElement();
                            }
                            catch
                            {
                                //if this is string value
                                this.nextWriter.WriteStartElement("x",EXCEL_NAMESPACE);
                                this.nextWriter.WriteStartAttribute("v");
                                value = pivotTable[row, col];
                                this.nextWriter.WriteString(this.fieldItems[col, 0][value].ToString());
                                this.nextWriter.WriteEndAttribute();
                                this.nextWriter.WriteEndElement();
                            }
                        }                         
                        this.nextWriter.WriteEndElement();
                    }
                    this.isInCacheRecords = false;
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

            this.fieldItems = new Hashtable[cols,2];

            for(int i = 0; i < 2; i++)
                for (int field = 0; field < cols; field++)
                    fieldItems[field,i] = new Hashtable();
            
            for (int row = 0; row < rows; row++)
                for (int col = 0; col < cols; col++)
                {
                    int keyCol = Convert.ToInt32(colStart) + col;
                    int keyRow = Convert.ToInt32(rowStart) + row;

                    string key = sheetNum + "#" + keyCol.ToString() + ":" + keyRow.ToString();

                    if (pivotCells.ContainsKey(key))
                    {                        
                        this.pivotTable[row, col] = (string)pivotCells[key];

                        if (!fieldItems[col,0].ContainsKey((string)pivotCells[key]))
                        {
                            fieldItems[col,0].Add((string)pivotCells[key], index[col]);
                            fieldItems[col,1].Add(index[col],(string)pivotCells[key]);
                            index[col]++;
                        }
                    }
                    else
                    {
                        this.pivotTable[row, col] = "";
                        if (!fieldItems[col,0].ContainsKey(""))
                        {
                            //fieldItems[col].Add("", index[col]);
                            fieldItems[col,0].Add("", index[col]);
                            fieldItems[col,1].Add(index[col],"");
                            index[col]++;
                        }
                    }
                }

            //Console.WriteLine(indeces[0].Count);
                //foreach (string key in this.fieldItems[0].Keys)
                //{
                //    Console.WriteLine("-" + fieldItems[0][key] + ":" + key);
                //}

        }
    }
}
