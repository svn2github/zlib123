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
using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using System.Collections;

using CleverAge.OdfConverter.OdfZipUtils;
using CleverAge.OdfConverter.OdfConverterLib;
using Tenuto.Reader;
using Tenuto.Grammar;
using Tenuto.Verifier;

namespace CleverAge.OdfConverter.OdfConverterTest
{
	/// <summary>Exception thrown when the file is not valid</summary>
	public class OdfValidatorException : Exception
	{
		public OdfValidatorException(String msg) : base (msg) {}
	}
	
	/// <summary>Check the validity of a odf file. Throw an OdfValidatorException if errors occurs</summary>
	public class OdfValidator
	{
        private const string RESOURCES_LOCATION = "resources";
		
        // namespaces and related schemas
        private static string ODF_SCHEMA = "odfschemas.OpenDocumentSchema10.rng";

        private XmlReader grammarName = null;
        private Grammar grammar = null;

		/// <summary>
		/// Initialize the validator
		/// </summary>
		public OdfValidator()
		{
            ResourceResolver resolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + RESOURCES_LOCATION);
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.None;
            settings.XmlResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + RESOURCES_LOCATION);
            XmlReader xmlReader = XmlReader.Create(ODF_SCHEMA, settings);
            this.grammarName = xmlReader;
        }
		
	    /// <summary>
	    /// Check the validity of an Office Open XML file.
	    /// </summary>
	    /// <param name="fileName">The path of the docx file.</param>
		public void validate(String fileName)
		{
            try 
            {
                this.grammar = new GrammarReader(new ConsoleController()).parse(this.grammarName);
            }
            catch (Exception e)
			{
				throw new OdfValidatorException("Problem parsing grammar file : " + e.Message);
			}
            ZipReader reader = null;
            try
            {
                reader = ZipFactory.OpenArchive(fileName);
                
            }
            catch (Exception e)
            {
                throw new OdfValidatorException("Problem opening the odt file : " + e.Message);
            }
            //content.xml
            //styles.xml
            //meta.xml
            //settings.xml
            try
            {
                Stream content = null;
                content = reader.GetEntry("content.xml");
                XmlReader xmlReader = XmlReader.Create(content);
                bool wasValid = Verifier.Verify(xmlReader, this.grammar, new ConsoleErrorReporter());
            }
            catch (ZipEntryNotFoundException e)
            {
                throw new OdfValidatorException("Entry not found in ODT file [content.xml]: " + e.Message);
            }
            catch (Exception e)
            {
                throw new OdfValidatorException("Problem validating ODT file [content.xml]: " + e.Message);
            }
            try
            {
                Stream content = null;
                content = reader.GetEntry("styles.xml");
                XmlReader xmlReader = XmlReader.Create(content);
                bool wasValid = Verifier.Verify(xmlReader, this.grammar, new ConsoleErrorReporter());
            }
            catch (ZipEntryNotFoundException e)
            {
                throw new OdfValidatorException("Entry not found in ODT file [styles.xml]: " + e.Message);
            }
            catch (Exception e)
            {
                throw new OdfValidatorException("Problem validating ODT file [styles.xml]: " + e.Message);
            }
            try
            {
                Stream content = null;
                content = reader.GetEntry("meta.xml");
                XmlReader xmlReader = XmlReader.Create(content);
                bool wasValid = Verifier.Verify(xmlReader, this.grammar, new ConsoleErrorReporter());
            }
            catch (ZipEntryNotFoundException e)
            {
                Console.WriteLine("Entry not found in ODT file [meta.xml]: " + e.Message);
            }
            catch (Exception e)
            {
                throw new OdfValidatorException("Problem validating ODT file [styles.xml]: " + e.Message);
            }
            try
            {
                Stream content = null;
                content = reader.GetEntry("settings.xml");
                XmlReader xmlReader = XmlReader.Create(content);
                bool wasValid = Verifier.Verify(xmlReader, this.grammar, new ConsoleErrorReporter());
            }
            catch (ZipEntryNotFoundException e)
            {
                Console.WriteLine("Entry not found in ODT file [settings.xml]: " + e.Message);
            }
            catch (Exception e)
            {
                throw new OdfValidatorException("Problem validating ODT file [styles.xml]: " + e.Message);
            }
        }
				
        public void ValidationHandler(object sender, ValidationEventArgs args)
      	{
        	throw new OdfValidatorException("XML Schema Validation error : " + args.Message);
      	}
	}
    
    public class ConsoleController : GrammarReaderController
    {
        public void error(string msg, IXmlLineInfo loc)
        {
            Console.WriteLine("Error: " + msg);
        }

        public void warning(string msg, IXmlLineInfo loc)
        {
            Console.WriteLine("Warning: " + msg);
        }
    }
    
    public class ConsoleErrorReporter : Tenuto.Verifier.ErrorHandler
    {
        public void Error(string msg)
        {
            Console.WriteLine(msg);
        }
    }
}
