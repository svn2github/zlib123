/* 
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
using System.Collections.Generic;
using System.Text;

namespace OdfConverter.Wordprocessing
{
    //public class OoxDefaultStyleVisibility
    //{
    //    private static Dictionary<string, OoxDefaultStyleVisibility> _properties;

    //    private bool _customStyle = false;

    //    private bool _hidden = false;
    //    private int _uiPriority = 99;
    //    private bool _semiHidden = false;
    //    private bool _unhideWhenUsed = false;
    //    private bool _qFormat = true;

    //    public bool CustomStyle
    //    {
    //        get { return _customStyle; }
    //        set { _customStyle = value; }
    //    }

    //    public bool Hidden
    //    {
    //        get { return _hidden; }
    //        set { _hidden = value; }
    //    }
        
    //    public int UiPriority
    //    {
    //        get { return _uiPriority; }
    //        set { _uiPriority = value; }
    //    }
        
    //    public bool SemiHidden
    //    {
    //        get { return _semiHidden; }
    //        set { _semiHidden = value; }
    //    }

    //    public bool UnhideWhenUsed
    //    {
    //        get { return _unhideWhenUsed; }
    //        set { _unhideWhenUsed = value; }
    //    }

    //    public bool QFormat
    //    {
    //        get { return _qFormat; }
    //        set { _qFormat = value; }
    //    }

    //    public OoxDefaultStyleVisibility(bool customStyle, bool hidden, int uiPriority, bool semiHidden, bool unhideWhenUsed, bool qFormat)
    //    {
    //        this.CustomStyle = customStyle;
    //        this.Hidden = hidden;
    //        this.UiPriority = uiPriority;
    //        this.SemiHidden = semiHidden;
    //        this.UnhideWhenUsed = unhideWhenUsed;
    //        this.QFormat = qFormat;
    //    }

    //    public static OoxDefaultStyleVisibility GetDefaultProperties(string styleId)
    //    {
    //        if (_properties == null)
    //        {
    //            _properties = new Dictionary<string, OoxDefaultStyleVisibility>();

    //            _properties.Add("1/1.1/1.1.1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("1/a/i", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("article/section", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("balloontext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bibliography", new OoxDefaultStyleVisibility(false, false, 37, true, true, false));
    //            _properties.Add("blocktext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytext2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytext3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytextfirstindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytextfirstindent2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytextindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytextindent2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("bodytextindent3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("booktitle", new OoxDefaultStyleVisibility(false, false, 33, false, false, true));
    //            _properties.Add("caption", new OoxDefaultStyleVisibility(false, false, 35, true, true, true));
    //            _properties.Add("closing", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("colorfulgrid", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent1", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent2", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent3", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent4", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent5", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfulgrid-accent6", new OoxDefaultStyleVisibility(false, false, 73, false, false, false));
    //            _properties.Add("colorfullist", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent1", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent2", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent3", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent4", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent5", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfullist-accent6", new OoxDefaultStyleVisibility(false, false, 72, false, false, false));
    //            _properties.Add("colorfulshading", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent1", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent2", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent3", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent4", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent5", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("colorfulshading-accent6", new OoxDefaultStyleVisibility(false, false, 71, false, false, false));
    //            _properties.Add("commentreference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("commentsubject", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("commenttext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("darklist", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent1", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent2", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent3", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent4", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent5", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("darklist-accent6", new OoxDefaultStyleVisibility(false, false, 70, false, false, false));
    //            _properties.Add("date", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("defaultparagraphfont", new OoxDefaultStyleVisibility(false, false, 1, true, true, false));
    //            _properties.Add("documentmap", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("e-mailsignature", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("emphasis", new OoxDefaultStyleVisibility(false, false, 20, false, false, true));
    //            _properties.Add("endnotereference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("endnotetext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("envelopeaddress", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("envelopereturn", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("followedhyperlink", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("footer", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("footnotereference", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("footnotetext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("header", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("heading1", new OoxDefaultStyleVisibility(false, false, 9, false, false, true));
    //            _properties.Add("heading2", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading3", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading4", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading5", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading6", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading7", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading8", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("heading9", new OoxDefaultStyleVisibility(false, false, 9, true, true, true));
    //            _properties.Add("htmlacronym", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmladdress", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlcite", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlcode", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmldefinition", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlkeyboard", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlpreformatted", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlsample", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmltypewriter", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("htmlvariable", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("hyperlink", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("index9", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("indexheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("intenseemphasis", new OoxDefaultStyleVisibility(false, false, 21, false, false, true));
    //            _properties.Add("intensequote", new OoxDefaultStyleVisibility(false, false, 30, false, false, true));
    //            _properties.Add("intensereference", new OoxDefaultStyleVisibility(false, false, 32, false, false, true));
    //            _properties.Add("lightgrid", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent1", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent2", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent3", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent4", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent5", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightgrid-accent6", new OoxDefaultStyleVisibility(false, false, 62, false, false, false));
    //            _properties.Add("lightlist", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent1", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent2", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent3", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent4", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent5", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightlist-accent6", new OoxDefaultStyleVisibility(false, false, 61, false, false, false));
    //            _properties.Add("lightshading", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent1", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent2", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent3", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent4", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent5", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("lightshading-accent6", new OoxDefaultStyleVisibility(false, false, 60, false, false, false));
    //            _properties.Add("linenumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("list", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("list2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("list3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("list4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("list5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listbullet", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listbullet2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listbullet3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listbullet4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listbullet5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listcontinue", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listcontinue2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listcontinue3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listcontinue4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listcontinue5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listnumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listnumber2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listnumber3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listnumber4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listnumber5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("listparagraph", new OoxDefaultStyleVisibility(false, false, 34, false, false, true));
    //            _properties.Add("macrotext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("mediumgrid1", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent1", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent2", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent3", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent4", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent5", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid1-accent6", new OoxDefaultStyleVisibility(false, false, 67, false, false, false));
    //            _properties.Add("mediumgrid2", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent1", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent2", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent3", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent4", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent5", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid2-accent6", new OoxDefaultStyleVisibility(false, false, 68, false, false, false));
    //            _properties.Add("mediumgrid3", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent1", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent2", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent3", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent4", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent5", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumgrid3-accent6", new OoxDefaultStyleVisibility(false, false, 69, false, false, false));
    //            _properties.Add("mediumlist1", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent1", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent2", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent3", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent4", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent5", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist1-accent6", new OoxDefaultStyleVisibility(false, false, 65, false, false, false));
    //            _properties.Add("mediumlist2", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent1", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent2", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent3", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent4", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent5", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumlist2-accent6", new OoxDefaultStyleVisibility(false, false, 66, false, false, false));
    //            _properties.Add("mediumshading1", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent1", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent2", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent3", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent4", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent5", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading1-accent6", new OoxDefaultStyleVisibility(false, false, 63, false, false, false));
    //            _properties.Add("mediumshading2", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent1", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent2", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent3", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent4", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent5", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("mediumshading2-accent6", new OoxDefaultStyleVisibility(false, false, 64, false, false, false));
    //            _properties.Add("messageheader", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("nolist", new OoxDefaultStyleVisibility(false, false, 99, false, false, false));
    //            _properties.Add("nospacing", new OoxDefaultStyleVisibility(false, false, 1, false, false, true));
    //            _properties.Add("normal", new OoxDefaultStyleVisibility(false, false, 0, false, false, true));
    //            _properties.Add("normal(web)", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("normalindent", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("noteheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("pagenumber", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("placeholdertext", new OoxDefaultStyleVisibility(false, true, 99, false, false, false));
    //            _properties.Add("plaintext", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("quote", new OoxDefaultStyleVisibility(false, false, 29, false, false, true));
    //            _properties.Add("salutation", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("signature", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("strong", new OoxDefaultStyleVisibility(false, false, 22, false, false, true));
    //            _properties.Add("subtitle", new OoxDefaultStyleVisibility(false, false, 11, false, false, true));
    //            _properties.Add("subtleemphasis", new OoxDefaultStyleVisibility(false, false, 19, false, false, true));
    //            _properties.Add("subtlereference", new OoxDefaultStyleVisibility(false, false, 31, false, false, true));
    //            _properties.Add("table3deffects1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("table3deffects2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("table3deffects3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableclassic1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableclassic2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableclassic3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableclassic4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolorful1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolorful2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolorful3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolumns1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolumns2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolumns3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolumns4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecolumns5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablecontemporary", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableelegant", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid", new OoxDefaultStyleVisibility(false, false, 59, false, false, false));
    //            _properties.Add("tablegrid1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablegrid8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist4", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist5", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist6", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist7", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablelist8", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablenormal", new OoxDefaultStyleVisibility(false, false, 99, false, false, false));
    //            _properties.Add("tableofauthorities", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableoffigures", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableprofessional", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablesimple1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablesimple2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablesimple3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablesubtle1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tablesubtle2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tabletheme", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableweb1", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableweb2", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("tableweb3", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("title", new OoxDefaultStyleVisibility(false, false, 10, false, false, true));
    //            _properties.Add("toaheading", new OoxDefaultStyleVisibility(false, false, 99, true, true, false));
    //            _properties.Add("toc1", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc2", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc3", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc4", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc5", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc6", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc7", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc8", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("toc9", new OoxDefaultStyleVisibility(false, false, 39, true, true, false));
    //            _properties.Add("tocheading", new OoxDefaultStyleVisibility(false, false, 39, true, true, true));
    //        }

    //        if (_properties.ContainsKey(styleId.Replace(" ", "").ToLower()))
    //        {
    //            return _properties[styleId.Replace(" ", "").ToLower()];
    //        }
    //        else if (styleId.ToLower().EndsWith("char") && _properties.ContainsKey(styleId.Substring(0, styleId.Length - 4).ToLower()))
    //        {
    //            // make Word-generated linked character styles invisible
    //            return new OoxDefaultStyleVisibility(true, false, 99, true, false, false);
    //        }

    //        return new OoxDefaultStyleVisibility(true, false, 0, false, false, true);
    //    }
    //}
}
