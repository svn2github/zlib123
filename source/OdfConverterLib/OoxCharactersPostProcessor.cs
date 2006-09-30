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

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for characters post processings
    public class OoxCharactersPostProcessor : AbstractPostProcessor
    {
        public OoxCharactersPostProcessor(XmlWriter nextWriter):base(nextWriter)
        {
        }

        public override void WriteString(string text)
        {
            this.ReplaceSoftHyphens(text);
        }

        private void ReplaceSoftHyphens(string text)
        {
            int i = 0;
            if ((i = text.IndexOf('\u00AD')) >= 0)
            {
                this.ReplaceNonBreakingHyphens(text.Substring(0, i));
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "softHyphen", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (i < text.Length - 1)
                {
                    nextWriter.WriteEndElement();
                    nextWriter.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    this.ReplaceSoftHyphens(text.Substring(i + 1, text.Length - i - 1));
                }
            }
            else
            {
                this.ReplaceNonBreakingHyphens(text);
            }
        }

        private void ReplaceNonBreakingHyphens(string text)
        {
            int i = 0;
            if ((i = text.IndexOf('\u2011')) >= 0)
            {
                nextWriter.WriteString(text.Substring(0, i));
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "noBreakHyphen", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (i < text.Length - 1)
                {
                    nextWriter.WriteEndElement();
                    nextWriter.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    this.ReplaceNonBreakingHyphens(text.Substring(i + 1, text.Length - i - 1));
                }
            }
            else
            {
                nextWriter.WriteString(text);
            }
        }
    }
}
