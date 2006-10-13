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

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// An <c>XmlWriter</c> implementation for table post processings
    public abstract class AbstractPostProcessor : XmlWriter
    {
        protected XmlWriter nextWriter;

        protected AbstractPostProcessor(XmlWriter nextWriter)
        {
            this.nextWriter = nextWriter;

            if (nextWriter == null)
            {
                throw new Exception("nextWriter can's be null");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (this.nextWriter != null)
            {
                this.nextWriter.Close();
            }
            this.nextWriter = null;
            base.Dispose(disposing);
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            this.nextWriter.WriteStartElement(prefix, localName, ns);
        }

        public override void WriteEndElement()
        {
            this.nextWriter.WriteEndElement();
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            this.nextWriter.WriteStartAttribute(prefix, localName, ns);
        }

        public override void WriteEndAttribute()
        {
            this.nextWriter.WriteEndAttribute();
        }

        public override void WriteString(string text)
        {
            this.nextWriter.WriteString(text);
        }

        public override void WriteFullEndElement()
        {
            this.nextWriter.WriteFullEndElement();
        }

        public override void WriteStartDocument()
        {
            this.nextWriter.WriteStartDocument();
        }

        public override void WriteStartDocument(bool b)
        {
            this.nextWriter.WriteStartDocument(b);
        }

        public override void WriteEndDocument()
        {
            this.nextWriter.WriteEndDocument();
        }

        public override void WriteDocType(string name, string pubid, string sysid, string subset)
        {
            this.nextWriter.WriteDocType(name, pubid, sysid, subset);
        }

        public override void WriteCData(string s)
        {
            this.nextWriter.WriteCData(s);
        }

        public override void WriteComment(string s)
        {
            this.nextWriter.WriteComment(s);
        }

        public override void WriteProcessingInstruction(string name, string text)
        {
            this.nextWriter.WriteProcessingInstruction(name, text);
        }

        public override void WriteEntityRef(string name)
        {
            this.nextWriter.WriteEntityRef(name);
        }

        public override void WriteCharEntity(char c)
        {
            this.nextWriter.WriteCharEntity(c);
        }

        public override void WriteWhitespace(string s)
        {
            this.nextWriter.WriteWhitespace(s);
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            this.nextWriter.WriteSurrogateCharEntity(lowChar, highChar);
        }

        public override void WriteChars(char[] buffer, int index, int count)
        {
            this.nextWriter.WriteChars(buffer, index, count);
        }

        public override void WriteRaw(char[] buffer, int index, int count)
        {
            this.nextWriter.WriteRaw(buffer, index, count);
        }

        public override void WriteRaw(string data)
        {
            this.nextWriter.WriteRaw(data);
        }

        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            this.nextWriter.WriteBase64(buffer, index, count);
        }

        public override WriteState WriteState
        {
            get
            {
                return this.nextWriter.WriteState;
            }
        }

        public override void Close()
        {
            this.nextWriter.Close();
        }

        public override void Flush()
        {
            this.nextWriter.Flush();
        }

        public override string LookupPrefix(string ns)
        {
            return this.nextWriter.LookupPrefix(ns);
        }

        /// <summary>
        /// Simple representation of elements or attributes nodes
        /// </summary>
        protected class Node
        {

            private string name;
            public string Name
            {
                get { return name; }
            }

            private string prefix;
            public string Prefix
            {
                get { return prefix; }
            }

            private string ns;
            public string Ns
            {
                get { return ns; }
            }

            public Node(string prefix, string name, string ns)
            {
                this.prefix = prefix;
                this.name = name;
                this.ns = ns;
            }

        }

        protected class Attribute : Node
        {
            private string value;

            public Attribute(string prefix, string name, string ns)
                : base(prefix, name, ns)
            {
            }
            
          	public Attribute(string prefix, string name, string val, string ns)
                : base(prefix, name, ns)
            {
              	this.value = val;
            }
              
            public string Value
            {
                get { return this.value; }
                set { this.value = value; }
            }

            public void Write(XmlWriter writer)
            {
                writer.WriteStartAttribute(this.Prefix, this.Name, this.Ns);
                writer.WriteString(this.value);
                writer.WriteEndAttribute();
            }
        }

        protected class Element : Node, IComparable
        {
            private ArrayList attributes;
            private ArrayList children;

            public Element(string prefix, string name, string ns)
                : base(prefix, name, ns)
            {
                this.attributes = new ArrayList();
                this.children = new ArrayList();
            }

            public int CompareTo(object obj)
            {
                if (obj is Element)
                {
                    Element el = (Element)obj;
                    return this.Name.CompareTo(el.Name);
                }
                else if (obj is string)
                {
                    string s = (String)obj;
                    return this.Name.CompareTo(s);
                }
                return 1;
            }

            public ArrayList Attributes
            {
                get { return this.attributes; }
            }

            public ArrayList Children
            {
                get { return this.children; }
            }

            public void AddAttribute(Attribute attribute)
            {
                this.attributes.Add(attribute);
            }

            public void AddChild(Element element)
            {
                this.children.Add(element);
            }

            public void AddFirstChild(Element element)
            {
                this.children.Insert(0, element);
            }

            public void AddChildIfNotExist(Element element)
            {
                if (this.GetChild(element.Name, element.Ns) == null)
                {
                    this.children.Add(element);
                }
            }

            public void AddChild(string text)
            {
                this.children.Add(text);
            }

            public void RemoveChild(Element element)
            {
                if (this.GetChild(element.Name, element.Ns) != null)
                {
                    this.children.Remove(element);
                }
            }

            public Element GetChild(string name, string ns)
            {
                foreach (object node in this.children)
                {
                    if (node is Element)
                    {
                        Element element = (Element)node;
                        if (element.Name.Equals(name) && element.Ns.Equals(ns))
                        {
                            return element;
                        }
                    }
                }
                return null;
            }

            public Attribute GetAttribute(string name, string ns)
            {
                foreach (Attribute attribute in this.attributes)
                {
                    if (attribute.Name.Equals(name) && attribute.Ns.Equals(ns))
                    {
                        return attribute;
                    }
                }
                return null;
            }

            public string GetAttributeValue(string name, string ns)
            {
                foreach (Attribute attribute in this.attributes)
                {
                    if (attribute.Name.Equals(name) && attribute.Ns.Equals(ns))
                    {
                        return attribute.Value;
                    }
                }
                return "";
            }

            public void SortChildren()
            {
                this.children.Sort();
            }

            public bool HasChild()
            {
                return this.children.Count > 0;
            }

            public void Write(XmlWriter writer)
            {
                writer.WriteStartElement(this.Prefix, this.Name, this.Ns);
                foreach (Attribute attribute in this.attributes)
                {
                    attribute.Write(writer);
                }
                foreach (object node in this.children)
                {
                    if (node is Element)
                    {
                        ((Element)node).Write(writer);
                    }
                    else if (node is string)
                    {
                        writer.WriteString((string)node);
                    }
                    else
                    {
                        throw new Exception("Child node must be of type Element or string");
                    }
                }
                writer.WriteEndElement();
            }

        }
    }
}
