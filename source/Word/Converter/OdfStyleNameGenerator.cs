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
using System.Text.RegularExpressions;

namespace OdfConverter.Wordprocessing
{
    //public class OdfStyleNameGenerator
    //{
    //    // see http://www.w3.org/TR/REC-xml-names/#NT-NCName for allowed characters in NCName
    //    private Regex _invalidLettersAtStart = new Regex(@"^([^\p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}_])(.*)", RegexOptions.Compiled); // valid start chars are \p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}_
    //    private Regex _invalidChars = new Regex(@"[^_\p{Ll}\p{Lu}\p{Lo}\p{Lt}\p{Nl}\p{Mc}\p{Me}\p{Mn}\p{Lm}\p{Nd}]", RegexOptions.Compiled);

    //    private Dictionary<string, string> _name2ncname = new Dictionary<string, string>();
    //    private Dictionary<string, string> _ncname2name = new Dictionary<string, string>();

    //    private static OdfStyleNameGenerator _instance;

    //    private OdfStyleNameGenerator()
    //    {
    //    }

    //    public static OdfStyleNameGenerator Instance
    //    {
    //        get
    //        {
    //            if (_instance == null)
    //            {
    //                _instance = new OdfStyleNameGenerator();
    //            }
    //            return _instance;
    //        }
    //    }

    //    /// <summary>
    //    /// Escapes all invalid characters from a style name and - if the name is not yet unique - 
    //    /// appends a counter to make the name unique
    //    /// </summary>
    //    public string NCNameFromString(string name)
    //    {
    //        string ncname = String.Empty;

    //        if (_name2ncname.ContainsKey(name))
    //        {
    //            ncname = _name2ncname[name];
    //        }
    //        else
    //        {
    //            // escape invalid characters
    //            ncname = name;
    //            Match invalidCharsMatch = _invalidChars.Match(ncname);
    //            foreach (Capture invalidCharCapture in invalidCharsMatch.Captures)
    //            {
    //                ncname = ncname.Replace(invalidCharCapture.Value, string.Format("_{0:x}_", (int)invalidCharCapture.Value[0]));
    //            }

    //            // escape invalid start character
    //            Match firstChar = _invalidLettersAtStart.Match(name);
    //            if (firstChar.Success)
    //            {
    //                ncname = string.Format("_{0:x}_{1}", (int)firstChar.Groups[1].Value[0], firstChar.Groups[2].Value);
    //            }

    //            // create new unique ncname
    //            int counter = 0;
    //            string uniqueName = ncname;
    //            while (_ncname2name.ContainsKey(uniqueName))
    //            {
    //                uniqueName = string.Format("{0}_{1}", ncname, counter);
    //                counter++;
    //            }
    //            ncname = uniqueName;
    //            _ncname2name.Add(ncname, name);
    //            _name2ncname.Add(name, ncname);
    //        }
    //        return ncname;
    //    }
    //}
}


