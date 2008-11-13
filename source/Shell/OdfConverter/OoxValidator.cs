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

namespace OdfConverter.CommandLineTool 
{
    interface IValidator 
    {
        void validate(string fileName);
    }

	/// <summary>Exception thrown when the file is not valid</summary>
	public class ValidationException : Exception
	{
		public ValidationException(String msg) : base (msg) {}
	}
	
	/// <summary>Check the validity of a docx file. Throw an OoxValidatorException if errors occurs</summary>
	public class OoxValidator : IValidator
	{
        private const string RESOURCES_LOCATION = "resources";

		// namespaces and related schemas
        private static string OOX_CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static string OOX_CONTENT_TYPE_SCHEMA = "ooxschemas/opc-contentTypeItem.xsd";
        private static string OOX_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static string OOX_RELATIONSHIP_SCHEMA = "ooxschemas/opc-relationshipPart.xsd";
        private static string OOX_META_CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static string OOX_META_CORE_SCHEMA = "ooxschemas/opc-coreProperties.xsd";
        private static string OOX_META_APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private static string OOX_META_APP_SCHEMA = "ooxschemas/shared-documentPropertiesExtended.xsd";
        private static string OOX_WORDML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static string OOX_WORDML_SCHEMA = "ooxschemas/wml.xsd";
        private static string OOX_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static string OOX_DML_STYLE_SCHEMA = "ooxschemas/dml-stylesheet.xsd";
		private static string OOX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		private static string OOX_WPDRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private static string OOX_DML_WPDRAWING_SCHEMA = "ooxschemas/dml-wordprocessingDrawing.xsd";
		private static string OOX_PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
		private static string OOX_DML_PICTURE_SCHEMA = "ooxschemas/dml-picture.xsd";	
		// OOX special files
        private static string OOX_CONTENT_TYPE_FILE = "[Content_Types].xml";
		private static string OOX_RELATIONSHIP_FILE = "_rels/.rels";

		// OOX relationship
        private static string OOX_DOCUMENT_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    
        private XmlReaderSettings settings = null; // global settings to open the xml files
        private ConversionReport report;
		
		/// <summary>
		/// Initialize the validator
		/// </summary>
		public OoxValidator(ConversionReport report)
		{
			this.settings = new XmlReaderSettings();
            this.report = report;
			
			// resolver
			EmbeddedResourceResolver resolver = new EmbeddedResourceResolver(Assembly.GetExecutingAssembly(), 
                this.GetType().Namespace, ".resources.", true);
			this.settings.XmlResolver = resolver;			
			
			// schemas
			this.settings.Schemas.XmlResolver = resolver;

			this.settings.Schemas.Add(OOX_RELATIONSHIP_NS, OOX_RELATIONSHIP_SCHEMA);
			this.settings.Schemas.Add(OOX_CONTENT_TYPE_NS, OOX_CONTENT_TYPE_SCHEMA);
			this.settings.Schemas.Add(OOX_META_CORE_NS, OOX_META_CORE_SCHEMA);
			this.settings.Schemas.Add(OOX_META_APP_NS, OOX_META_APP_SCHEMA);
			this.settings.Schemas.Add(OOX_WORDML_NS, OOX_WORDML_SCHEMA);
			this.settings.Schemas.Add(OOX_DML_NS, OOX_DML_STYLE_SCHEMA);
			this.settings.Schemas.Add(OOX_PICTURE_NS, OOX_DML_PICTURE_SCHEMA);
			this.settings.Schemas.Add(OOX_WPDRAWING_NS, OOX_DML_WPDRAWING_SCHEMA);

			this.settings.ValidationType = ValidationType.Schema;
			this.settings.ValidationEventHandler += new ValidationEventHandler (ValidationHandler);
		}
	
		/// <summary>
	    /// Check the validity of an Office Open XML file.
	    /// </summary>
	    /// <param name="fileName">The path of the docx file.</param>
		public void validate(string fileName)
		{
			// 0. The file must exist and be a valid Zip archive
			ZipReader reader = null;
			try
			{
				reader = ZipFactory.OpenArchive(fileName);
			}
			catch (Exception e)
			{
				throw new ValidationException("Problem opening the docx file : " + e.Message);
			}
			
			// 1. [Content_Types].xml must be present and valid
			Stream contentTypes = null;
			try
			{
				contentTypes = reader.GetEntry(OOX_CONTENT_TYPE_FILE);
			}
			catch (Exception)
			{
				throw new ValidationException("The docx package must have a \"/[Content_Types].xml\" file");
			}
			this.validateXml(contentTypes);

			// 2. _rels/.rels must be present and valid
			Stream relationShips = null;
			try
			{
				relationShips = reader.GetEntry(OOX_RELATIONSHIP_FILE);
			}
			catch (Exception)
			{
				throw new ValidationException("The docx package must have a \"/_rels/.rels\" file");
			}
			this.validateXml(relationShips);
			
			// 3. _rel/.rels must contain a relationship of type openDocument
			relationShips = reader.GetEntry(OOX_RELATIONSHIP_FILE);
			XmlReader r = XmlReader.Create(relationShips);
			String docTarget = null;
			while (r.Read() && docTarget == null)
			{
				if (r.NodeType == XmlNodeType.Element && r.GetAttribute("Type") == OOX_DOCUMENT_RELATIONSHIP_TYPE)
				{
					docTarget = r.GetAttribute("Target");
				}
			}
			if (docTarget == null)
			{
				throw new ValidationException("openDocument relation not found in \"/_rels/.rels\"");
			}
			
			// 4. For each item in _rels/.rels
			relationShips = reader.GetEntry(OOX_RELATIONSHIP_FILE);
			r = XmlReader.Create(relationShips);
			while (r.Read())
			{
				if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
				{
					String target = r.GetAttribute("Target");
					
					
					
					// 4.1. The target item must exist in the package
					Stream item = null;
					try
					{
						item = reader.GetEntry(target);
					}
					catch (Exception)
					{
						throw new ValidationException("The file \"" + target + "\" is described in the \"/_rels/.rels\" file but does not exist in the package.");
					}
					
					// 4.2. A content type can be found in [Content_Types].xml file
					String ct = this.findContentType(reader, "/" + target);
					
					// 4.3. If it's an xml file, it has to be valid
					if (ct.EndsWith("+xml"))
					{
						this.validateXml(item);
					}
				}
			}
			
			// Does a part relationship exist ?
			Stream partRel = null;
			String partDir = docTarget.Substring(0, docTarget.IndexOf("/"));
			String partRelPath = partDir + "/_rels/" + docTarget.Substring(docTarget.IndexOf("/") + 1) + ".rels";
			bool partRelExists = true;
			try
			{
				partRel = reader.GetEntry(partRelPath);
			}
			catch ( Exception )
			{
				partRelExists = false;
			}
			
			if ( partRelExists )
			{
				this.validateXml(partRel);
				
				// 5. For each item in /word/_rels/document.xml.rels
				partRel = reader.GetEntry(partRelPath);
				r = XmlReader.Create(partRel);
				while (r.Read())
				{
					if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
					{
						String target = partDir + "/" + r.GetAttribute("Target");

						// Is the target item exist in the package ?
						Stream item = null;
						bool fileExists = true;
						try
						{
							item = reader.GetEntry(target);
						}
						catch (Exception)
						{
							//throw new OoxValidatorException("The file \"" + target + "\" is described in the \"/word/_rels/document.xml.rels\" file but does not exist in the package.");
							fileExists = false;
						}
						
						if (fileExists)
						{
							// 5.1. A content type can be found in [Content_Types].xml file
							String ct = this.findContentType(reader, "/" + target);
						
							// 5.2. If it's an xml file, it has to be valid
							if (ct.EndsWith("+xml"))
							{
								this.validateXml(item);
							}
						}
					}
				}
			}
			
			// 6. do all referenced relationships exist in the part relationship file
			
			// retrieve all ids referenced in the document
			
			Stream doc = reader.GetEntry(docTarget);
			r = XmlReader.Create(doc);
			ArrayList ids = new ArrayList();
			while (r.Read())
			{
				if (r.NodeType == XmlNodeType.Element && r.GetAttribute("id", OOX_DOC_REL_NS) != null)
				{
					ids.Add(r.GetAttribute("id", OOX_DOC_REL_NS));
				}
			}
			
			// check if each id exists in the partRel file
			
			if (ids.Count != 0)
			{
				if (!partRelExists)
				{
					throw new ValidationException("Referenced id exist but no part relationship file found");
				}
				relationShips = reader.GetEntry(partRelPath);
				r = XmlReader.Create(relationShips);
				while (r.Read())
				{
					if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
					{
						if (ids.Contains(r.GetAttribute("Id")))
						{
							ids.Remove(r.GetAttribute("Id"));
						}
					}
				}
				if (ids.Count != 0)
				{
					throw new ValidationException("One or more relationship id have not been found in the partRelationship file : " + ids[0]);
				}
			}
		}
		
		// validate xml stream
		private void validateXml(Stream xmlStream)
		{
			XmlReader r = XmlReader.Create(xmlStream, this.settings);
			while(r.Read());
		}
		
		// find the content type of a part in the package
		private String findContentType(ZipReader reader, String target)
		{
			String extension = null;
			if (target.IndexOf(".") != -1)
			{
				extension = target.Substring(target.IndexOf(".") + 1);
			}
			Stream contentTypes = reader.GetEntry(OOX_CONTENT_TYPE_FILE);
			XmlReader r = XmlReader.Create(contentTypes);
			bool overrided = false;
			String contentType = null;
			while (r.Read())
			{
				if (r.NodeType == XmlNodeType.Element)
				{
					if (r.LocalName == "Default" && extension != null && r.GetAttribute("Extension") == extension && !overrided)
					{
						contentType = r.GetAttribute("ContentType");
					} else if (r.LocalName == "Override" && r.GetAttribute("PartName") == target)
					{
						overrided = true;
						contentType = r.GetAttribute("ContentType");
					}
				}
			}
			if ( contentType == null )
			{
				throw new ValidationException("Content type not found for " + target);
			}
			return contentType;
		}
		
        public void ValidationHandler(object sender, ValidationEventArgs args)
      	{
        	throw new ValidationException("XML Schema Validation error : " + args.Message);
      	}
	}
}
