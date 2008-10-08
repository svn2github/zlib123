using System;
using System.Collections.Generic;
using System.Text;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Xml;

namespace OdfConverter.Wordprocessing
{
    public class OoxReplacementPostProcessor : AbstractPostProcessor
    {
        private static bool _isBookMarkElement;
        private static bool _isBoorkmarkId;
        
        public OoxReplacementPostProcessor(XmlWriter nextWriter)
            : base(nextWriter)
		{
		}

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            //This is a bookmark element
            if (prefix == "w" && (localName == "bookmarkStart" || localName == "bookmarkEnd"))
            {
                _isBookMarkElement = true;
            }
            else
            {
                _isBookMarkElement = false;
            }

            this.nextWriter.WriteStartElement(prefix, localName, ns);
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            //this is a bookmark ID attribute
            if (_isBookMarkElement && prefix == "w" && localName == "id")
            {
                _isBoorkmarkId = true;
            }
            else
            {
                _isBoorkmarkId = false;
            }

            this.nextWriter.WriteStartAttribute(prefix, localName, ns);
        }

        public override void WriteString(string text)
        {
            string replacement = text;

            if (_isBookMarkElement && _isBoorkmarkId && text.StartsWith("http://www.dialogika.de/replace/bookmarkid/"))
            {
                replacement = "";

                //Replace bookmark IDs
                char[] id = text.Substring(text.LastIndexOf("/") + 1).ToLower().ToCharArray();

                //Replace the chars by their numbers
                for (int i = 0; i < id.Length; i++)
                {
                    replacement += (int)id[i];
                }
            }

            this.nextWriter.WriteString(replacement);
        }
    }
}
