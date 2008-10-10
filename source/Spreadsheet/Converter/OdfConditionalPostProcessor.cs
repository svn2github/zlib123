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
RefNo-1 08-sep-2008 Sandeep S     New feature   Changes for formula implementation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
*/

using System;
using System.Xml;
using System.Collections;
using CleverAge.OdfConverter.OdfConverterLib;
using CleverAge.OdfConverter.OdfZipUtils;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Windows.Forms;



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
                
                this.nextWriter.WriteString(ConditionalModification(stringSqref[text].ToString()));              
                ElementName = "";
            }
            //Start of RefNo-1
            else if (text.StartsWith("sonataOdfFormula"))
            {
                this.nextWriter.WriteString(GetFormula(text.Substring(16)));
            }
            else if (text.StartsWith("sonataColumnWidth:"))
            {
                this.nextWriter.WriteString(GetColumnWidth(text.Substring(18)));
            }
            
            //End of RefNo-1
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


        protected int GetRowId(string value)
        {
            int result = 0;

            foreach (char c in value)
            {
                if (c >= '0' && c <= '9')
                {
                    result = (10 * result) + (c - '0');
                }
            }

            return result;
        }

        protected int GetColId(string value)
        {
            int result = 0;

            foreach (char c in value)
            {
                if (c >= 'A' && c <= 'Z')
                {
                    result = (26 * result) + (c - 'A' + 1);
                }
                else
                {
                    break;
                }
            }

            return result;
        }

        protected string ConditionalModification(string value)
        {
            string result = "";
            string prevCell = "";
            int prevRowNum = -1;
            int prevColNum = -1;
            int repeat = 0;

            int thisRowNum = -1;
            int thisColNum = -1;


            //result = value;
            foreach (string single in value.Split())
            {
                if (single.Contains(":"))
                {
                    if (result != "")
                    {
                        if (repeat > 0)
                        {
                            result = result + ":" + prevCell + " " + single;
                            prevCell = "";
                            prevRowNum = -1;
                            prevColNum = -1;
                            repeat = 0;
                        }
                        else
                        {
                            result = result + " " + single;
                        }

                    }
                    else
                    {
                        result = single;
                    }
                }
                else
                {
                    if (result != "")
                    {
                        thisRowNum = GetRowId(single);
                        thisColNum = GetColId(single);
                        if (prevCell != "")
                        {
                            if (thisColNum == prevColNum && thisRowNum == (prevRowNum + 1))
                            {
                                prevCell = single;
                                prevRowNum = thisRowNum;
                                prevColNum = thisColNum;
                                repeat = repeat + 1;
                            }
                            else if (repeat > 0)
                            {
                                result = result + ":" + prevCell + " " + single;
                                prevCell = single;
                                prevRowNum = thisRowNum;
                                prevColNum = thisColNum;
								repeat = 0;
                            }
                            else
                            {
                                result = result + " " + single;
                                prevCell = single;
                                prevRowNum = thisRowNum;
                                prevColNum = thisColNum;
                            }
                        }
                    }
                    else
                    {
                        result = single;
                        prevCell = single;
                        prevRowNum = thisRowNum;
                        prevColNum = thisColNum;
                    }
                }
            }

            if (repeat > 0 && prevCell != "")
            {
                result = result + ":" + prevCell;
            }
            

            return result;
        }

        //Start of RefNo-1 : Methods for formula translation
        protected string GetFormula(string strOdfFormula)
        {
            string strOoxFormula = "";
            if (strOdfFormula != "")
            {
                //TODO : chk for the other namespaces:Eg Engginering formulas.
                if (strOdfFormula.StartsWith("oooc:"))
                {
                    strOoxFormula = TranslateToOoxFormula(strOdfFormula.Substring(6));
                }
                else
                {
                    strOoxFormula = strOdfFormula;
                }
            }
            else
            {
                strOoxFormula = strOdfFormula;
            }
            return strOoxFormula;
        }

        protected string TranslateToOoxFormula(string strOdfFormula)
        {
            string strFinalFormula = "";
            string strExpToTrans = "";

            //TODO : chk for other representations starting with &_;
            //to replace &quot; with " operator.
            strOdfFormula = strOdfFormula.Replace("&quot;", "\"");
            //to replace parameter seperation operator.
            strOdfFormula = strOdfFormula.Replace(';', ',');
            //to replace union operator 
            strOdfFormula = strOdfFormula.Replace('!', ' ');
            //replace &apos; with ' operator
            strOdfFormula = strOdfFormula.Replace("&apos;", "'");
            /*The functions whose names end with _ADD return the same results as the corresponding Microsoft Excel functions. Use the functions without _ADD to get results based on international standards. For example, the WEEKNUM function calculates the week number of a given date based on international standard ISO 6801, while WEEKNUM_ADD returns the same week number as Microsoft Excel.
             XML representation contains the namespace com.sun.star.sheet.addin.Analysis.getWeeknum(*/
            strOdfFormula = strOdfFormula.Replace("com.sun.star.sheet.addin.Analysis.get", "");
            strOdfFormula = strOdfFormula.Replace("com.sun.star.sheet.addin.DateFunctions.get", "");
            strOdfFormula = strOdfFormula.Replace("com.sun.star.sheet.addin.DateFunctions.getDiff", "");

            strOdfFormula = strOdfFormula.Replace("ERRORTYPE(", "ERROR.TYPE(");
            //chk for parameter discripency

            //TODO : chk for the '[' within the string
            //to translate cell reference
            if (strOdfFormula.Contains("["))
            {
                while (strOdfFormula.Contains("["))
                {
                    int intStart = strOdfFormula.IndexOf('[');
                    int intEnd = strOdfFormula.IndexOf(']');
                    strFinalFormula = strFinalFormula + strOdfFormula.Substring(0, intStart);
                    strExpToTrans = strOdfFormula.Substring(intStart + 1, intEnd - intStart - 1);
                    strOdfFormula = strOdfFormula.Substring(intEnd + 1);

                    strFinalFormula = strFinalFormula + TranslateToOoxCellRef(strExpToTrans);
                    if (!strOdfFormula.Contains("["))
                    {
                        strFinalFormula = strFinalFormula + strOdfFormula;
                    }
                }
            }
            else
            {
                strFinalFormula = strOdfFormula;
            }

            return strFinalFormula = TransFormulaParms(strFinalFormula);

        }

        protected string TranslateToOoxCellRef(string strOdfCellRef)
        {
            if (strOdfCellRef.StartsWith("."))
            {
                return strOdfCellRef = strOdfCellRef.Replace(".", ""); 
            }
            else
            {
                //TODO : chk for sheet name
                strOdfCellRef = strOdfCellRef.Replace(":.", ":");
                //return strOdfCellRef = strOdfCellRef.Replace('.', '!');
                
                strOdfCellRef = strOdfCellRef.Replace('.', '!');
                if (strOdfCellRef.Substring(0, strOdfCellRef.IndexOf('!')).StartsWith("$"))
                {
                    strOdfCellRef = strOdfCellRef.Replace("$", "");
                }
                if (!strOdfCellRef.Substring(0, strOdfCellRef.IndexOf('!')).StartsWith("'"))
                {
                    strOdfCellRef = "'" + strOdfCellRef.Replace("!", "'!");
                }                
                
                return strOdfCellRef;
            }
        }

        protected string TransFormulaParms(string strFormula)
        {
            string strTransFormula = "";
            string strFormulaToTrans = "";
            string strExpression = "";

            //chking for key word
            if (strFormula.Contains("ADDRESS(") || strFormula.Contains("CELING(") || strFormula.Contains("FLOOR("))
            {
                while (strFormula != "")
                {
                    //TODO : Include other formulas. need to chk which formula is comming first.
                    if (strFormula.Contains("ADDRESS(") || strFormula.Contains("CELING(") || strFormula.Contains("FLOOR("))
                    {
                        int[] arrIndex = new int[3];

                        string strFunction = "";
                        int intFunction = 0;
                        int intIndex = 0;

                        arrIndex[0] = strFormula.IndexOf("ADDRESS(");//ADD 4th parm:1
                        arrIndex[1] = strFormula.IndexOf("CELING(");//remove 3rd parm
                        arrIndex[2] = strFormula.IndexOf("FLOOR(");//remove 3rd parm


                        if (arrIndex[0] >= 0 && arrIndex[0] > arrIndex[1] && arrIndex[0] > arrIndex[2])
                        {
                            strFunction = "ADDRESS(";
                            intFunction = 1;
                            intIndex = arrIndex[0];

                        }
                        else if (arrIndex[1] >= 0 && arrIndex[1] > arrIndex[0] && arrIndex[1] > arrIndex[2])
                        {
                            strFunction = "CELING(";
                            intFunction = 2;
                            intIndex = arrIndex[1];
                        }
                        else if (arrIndex[2] >= 0 && arrIndex[2] > arrIndex[1] && arrIndex[2] > arrIndex[0])
                        {
                            strFunction = "FLOOR(";
                            intFunction = 3;
                            intIndex = arrIndex[2];
                        }
                        //|| strFormula.Contains("CELING(") || strFormula.Contains("FLOOR("))

                        string strConvertedExp = "";
                        strTransFormula = strTransFormula + strFormula.Substring(0, intIndex);
                        strFormulaToTrans = strFormula.Substring(intIndex);
                        if (strFormulaToTrans.Contains(")"))
                        {
                            strExpression = GetExpression(strFormulaToTrans);
                        }
                        if (strExpression != "")
                        {
                            //TODO : if function with one parm and second to be added.
                            if (strExpression.Contains(","))
                            {
                                if (intFunction == 1)
                                {
                                    strConvertedExp = AddRemoveParm(strExpression, 4, true, false, "1");
                                }
                                else
                                {
                                    strConvertedExp = AddRemoveParm(strExpression, 3, false, true, "");
                                }
                            }
                            else
                            {
                                strConvertedExp = strExpression;
                            }
                            //use converted and send remaining for chk
                            strTransFormula = strTransFormula + strConvertedExp;
                            strFormula = strFormulaToTrans.Remove(0, strExpression.Length);
                        }
                        else
                        {
                            //remove the key word and send remaing for chk
                            strTransFormula = strTransFormula + strFormulaToTrans.Substring(0, strFunction.Length);
                            strFormula = strFormulaToTrans.Substring(strFormula.Length);
                        }
                    }
                    else
                    {
                        strTransFormula = strTransFormula + strFormula;
                        strFormula = "";
                    }
                }
            }
            else
            {
                strTransFormula = strFormula;
            }
            return strTransFormula;
        }

        protected string AddRemoveParm(string strExpresion, int intPossition, bool blnParmAdd, bool blnParmRemove, string strParmAdd)
        {
            string strTransFormula = "";
            string strFormulaKeyword = strExpresion.Substring(0, strExpresion.IndexOf('('));
            strExpresion = strExpresion.Substring(strExpresion.IndexOf('(') + 1);
            strExpresion = strExpresion.Substring(0, strExpresion.Length - 1);

            ArrayList arlParms = GetParms(strExpresion);

            if (arlParms.Count >= intPossition)
            {
                if (blnParmAdd == true && blnParmRemove == false)
                {
                    strTransFormula = strFormulaKeyword + "(";
                    for (int i = 0; i < arlParms.Count; i++)
                    {
                        if (intPossition == i + 1)
                        {
                            strTransFormula = strTransFormula + strParmAdd + "," + arlParms[i].ToString() + ",";
                        }
                        else
                        {
                            strTransFormula = strTransFormula + arlParms[i].ToString() + ",";
                        }
                    }
                    strTransFormula = strTransFormula.Substring(0, strTransFormula.Length - 1) + ")";
                }
                else if (blnParmAdd == false && blnParmRemove == true)
                {
                    strTransFormula = strFormulaKeyword + "(";
                    for (int i = 0; i < arlParms.Count; i++)
                    {
                        if (intPossition != i + 1)
                        {
                            strTransFormula = strTransFormula + arlParms[i].ToString() + ",";
                        }
                    }
                    strTransFormula = strTransFormula.Substring(0, strTransFormula.Length - 1) + ")";
                }
            }
            else
            {
                strTransFormula = strFormulaKeyword + "(" + strExpresion + ")";
            }
            return strTransFormula;

        }

        protected string GetExpression(string strExpression)
        {
            string strChkValidExp = "";
            bool blnValidExp = false;
            int intChkFrom = 0;

            while (!blnValidExp)
            {
                if (intChkFrom == 0)
                    intChkFrom = strExpression.IndexOf(')');
                else
                    intChkFrom = strExpression.IndexOf(')', intChkFrom + 1);
                strChkValidExp = strExpression.Substring(0, intChkFrom + 1);

                blnValidExp = IsValidExpression(strChkValidExp);
            }

            return strChkValidExp;


        }

        protected ArrayList GetParms(string strExpression)
        {
            ArrayList arlParms = new ArrayList();

            string[] strArrParm = strExpression.Split(',');
            string strExpToVald = "";
            if (strArrParm.Length > 0)
            {

                for (int i = 0; i < strArrParm.Length; i++)
                {
                    if (strExpToVald == "")
                    {
                        strExpToVald = strArrParm[i].ToString();
                    }
                    else
                    {
                        strExpToVald = strExpToVald + "," + strArrParm[i].ToString();
                    }
                    if (strExpToVald.StartsWith("\""))
                    {
                        if (IsValidString(strExpToVald))
                        {
                            arlParms.Add(strExpToVald);
                            strExpToVald = "";
                        }
                    }
                    else if (IsValidExpression(strExpToVald))
                    {
                        strExpToVald = TransFormulaParms(strExpToVald);
                        arlParms.Add(strExpToVald);
                        strExpToVald = "";
                    }
                }
            }

            return arlParms;

        }

        protected bool IsValidString(string strExpression)
        {
            int intQoutes;
            MatchCollection matches = Regex.Matches(strExpression, @"[""]");
            intQoutes = matches.Count;
            intQoutes = intQoutes % 2;

            if (intQoutes == 0)
                return true;
            else
                return false;
        }

        protected bool IsValidExpression(string strExpression)
        {
            int intOpenBracket;
            int intClosedBracket;
            int intOpenFlower;
            int intClosedFlower;
            int intQoutes;


            MatchCollection matches = Regex.Matches(strExpression, @"[""]");
            intQoutes = matches.Count;
            intQoutes = intQoutes % 2;

            if (intQoutes == 0)
            {
                string strChkQoutes = strExpression;
                while (strChkQoutes.Contains(@""""))
                {
                    int intStart = 0;
                    int intEnd = 0;
                    intStart = strChkQoutes.IndexOf('"');
                    intEnd = strChkQoutes.IndexOf('"', intStart + 1);

                    string strRemvChars = strChkQoutes.Substring(intStart + 1, intEnd - intStart - 1);

                    Regex r = new Regex(@"[(){}]");
                    strRemvChars = r.Replace(strRemvChars, " ");

                    strChkQoutes = strChkQoutes.Substring(0, intStart) + strRemvChars + strChkQoutes.Substring(intEnd + 1);
                    //strChkQoutes = strChkQoutes.Substring(0, intStart);
                }
                strExpression = strChkQoutes;
            }

            matches = Regex.Matches(strExpression, @"[/(]");
            intOpenBracket = matches.Count;
            matches = Regex.Matches(strExpression, @"[/)]");
            intClosedBracket = matches.Count;
            matches = Regex.Matches(strExpression, @"[/{]");
            intOpenFlower = matches.Count;
            matches = Regex.Matches(strExpression, @"[/}]");
            intClosedFlower = matches.Count;
            matches = Regex.Matches(strExpression, @"[""]");
            intQoutes = matches.Count;
            intQoutes = intQoutes % 2;

            if (intOpenBracket == intClosedBracket && intOpenFlower == intClosedFlower && intQoutes == 0)
            {
                return true;
            }
            else
            {
                return false;
            }


        }
        //End of RefNo-1

        protected string GetColumnWidth(string text)
        {

            string measureString = "0";
            Font stringFont = new Font(text.Split('|')[0].ToString(), float.Parse(text.Split('|')[1].ToString()));
            Form dummyForm = new Form();
            Graphics g = dummyForm.CreateGraphics();
            System.Drawing.SizeF size = g.MeasureString(measureString, stringFont);
            //RectangleF layoutRect = new RectangleF(0, 0, size.Width, size.Height);
            RectangleF layoutRect = new RectangleF(50.0F, 50.0F, size.Width, size.Height);
            // Set string format.
            CharacterRange[] characterRanges = { new CharacterRange(0, 1) };
            StringFormat stringFormat = new StringFormat();
            stringFormat.SetMeasurableCharacterRanges(characterRanges);
            Region[] result = g.MeasureCharacterRanges(measureString, stringFont, layoutRect, stringFormat);
            //charWidth is nothing but  
            double maxDigWidth = result[0].GetBounds(g).Width - 2; // minus 2 bound borders
            //width=Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            //Column Width =Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
            dummyForm.Close();
            double columnWidth = System.Math.Truncate((((8 * maxDigWidth) + 5) / maxDigWidth) * 256) / 256;
            double columnWidthPx = System.Math.Truncate((((256 * columnWidth) + (System.Math.Truncate(128 / maxDigWidth))) / 256) * maxDigWidth);
            ////double columnWidthPx = System.Math.Truncate(((256 * columnWidth) + System.Math.Truncate(128 / maxDigWidth) / 256) * maxDigWidth);
            //double columnWidthInch = (columnWidthPx / 96.21212);
            double columnWidthPoints = columnWidthPx / 96.21212 * 72;
            return columnWidthPoints.ToString();

        }
    }
}
