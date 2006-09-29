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
using System.Collections;
using System.IO;
using System.Xml;
using System.Text;
using System.Diagnostics;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for table post processings
    public class OoxCharactersPostProcessor : XmlWriter
    {
        private XmlWriter nextWriter;

        /// <summary>
        /// Constructor
        /// </summary>
        public OoxCharactersPostProcessor(XmlWriter nextWriter)
        {
            this.nextWriter = nextWriter;

            if (nextWriter == null)
            {
                throw new Exception("Error in OoxTablePostProcessor contructor: null writer");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose managed resources
                if (nextWriter != null)
                    nextWriter.Close();
            }
            base.Dispose(disposing);

        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
           nextWriter.WriteStartElement(prefix, localName, ns);
        }

        public override void WriteEndElement()
        {
            nextWriter.WriteEndElement();
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            nextWriter.WriteStartAttribute(prefix, localName, ns);
        }

        public override void WriteEndAttribute()
        {
            nextWriter.WriteEndAttribute();
        }

        public override void WriteString(string text)
        {
            this.ReplaceSoftHyphens(text);
        }

        public override void WriteFullEndElement()
        {
            this.WriteFullEndElement();
        }

        public override void WriteStartDocument()
        {
            // nothing to do here
        }

        public override void WriteStartDocument(bool b)
        {
            // nothing to do here
        }

        public override void WriteEndDocument()
        {
            // nothing to do here
        }

        public override void WriteDocType(string name, string pubid, string sysid, string subset)
        {
            // nothing to do here
        }

        public override void WriteCData(string s)
        {
            nextWriter.WriteCData(s);
        }

        public override void WriteComment(string s)
        {
            nextWriter.WriteComment(s);
        }

        public override void WriteProcessingInstruction(string name, string text)
        {
            nextWriter.WriteProcessingInstruction(name, text);
        }

        public override void WriteEntityRef(String name)
        {
            nextWriter.WriteEntityRef(name);
        }

        public override void WriteCharEntity(char c)
        {
            nextWriter.WriteCharEntity(c);
        }

        public override void WriteWhitespace(string s)
        {
            nextWriter.WriteWhitespace(s);
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            nextWriter.WriteSurrogateCharEntity(lowChar, highChar);
        }

        public override void WriteChars(char[] buffer, int index, int count)
        {
            nextWriter.WriteChars(buffer, index, count);
        }

        public override void WriteRaw(char[] buffer, int index, int count)
        {
            nextWriter.WriteRaw(buffer, index, count);
        }

        public override void WriteRaw(string data)
        {
            nextWriter.WriteRaw(data);
        }

        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            nextWriter.WriteBase64(buffer, index, count);
        }

        public override WriteState WriteState
        {
            // nothing smart to do here
            get
            {
                return nextWriter.WriteState;
            }
        }

        public override void Close()
        {
            nextWriter.Close();
            nextWriter = null;

        }

        public override void Flush()
        {
            // nothing to do
        }

        public override string LookupPrefix(String ns)
        {
            return nextWriter.LookupPrefix(ns);
        }

        public void ReplaceSoftHyphens(string text)
        {
            int i = 0;
            if ((i = text.IndexOf('\u00AD')) >= 0)
            {
                this.ReplaceNonBreakingHyphens(text.Substring(0, i));
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "softHyphen", "http://schemas.openxmlformats.org/wordprocessingml/2006/3/main");
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/3/main");
                this.ReplaceSoftHyphens(text.Substring(i + 1, text.Length - i - 1));
            }
            else
            {
                this.ReplaceNonBreakingHyphens(text);
            }
        }

        public void ReplaceNonBreakingHyphens(string text)
        {
            int i = 0;
            if ((i = text.IndexOf('\u2011')) >= 0)
            {
                nextWriter.WriteString(text.Substring(0, i));
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "noBreakHyphen", "http://schemas.openxmlformats.org/wordprocessingml/2006/3/main");
                nextWriter.WriteEndElement();
                nextWriter.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/3/main");
                this.ReplaceNonBreakingHyphens(text.Substring(i + 1, text.Length - i - 1));
            }
            else
            {
                nextWriter.WriteString(text);
            }
        }
    }
}
