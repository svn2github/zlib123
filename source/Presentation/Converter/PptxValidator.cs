/* 
 * Copyright (c) 2006, Sonata Software Limited
 * All rights reserved.
 * 
 
 */

using System;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using System.Collections;
using CleverAge.OdfConverter.OdfZipUtils; 
using CleverAge.OdfConverter.OdfConverterLib;

namespace Sonata.OdfConverter.Presentation
{
    /// <summary>Exception thrown when the file is not valid</summary>
    public class PptxValidatorException : Exception
    {
        public PptxValidatorException(String msg) : base(msg) { }
    }

    /// <summary>Check the validity of a pptx file. Throw an PptxValidatorException if errors occurs</summary>
    public class PptxValidator
    {
        private const string RESOURCES_LOCATION = "resources";
        private ZipReader reader = null;

        // namespaces and related schemas
        private static string OOX_CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static string OOX_CONTENT_TYPE_SCHEMA = "ooxschemas/opc-contentTypeItem.xsd";
        private static string OOX_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static string OOX_RELATIONSHIP_SCHEMA = "ooxschemas/opc-relationshipPart.xsd";
        private static string OOX_META_CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private static string OOX_META_CORE_SCHEMA = "ooxschemas/opc-coreProperties.xsd";
        private static string OOX_META_APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private static string OOX_META_APP_SCHEMA = "ooxschemas/shared-documentPropertiesExtended.xsd";
         
        // OOX special files
        private static string OOX_CONTENT_TYPE_FILE = "[Content_Types].xml";
        private static string OOX_RELATIONSHIP_FILE = "_rels/.rels";
        private static string OOX_SIDEMASTER_RELATIONSHIP_FILE = "ppt/slideMasters";
        
        
        // OOX relationship
        private static string OOX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static string OOX_DOCUMENT_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private static string OOX_SIDEMASTER_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
        
        private XmlReaderSettings settings = null; // global settings to open the xml files
        //private Report report;

        /// <summary>
        /// Initialize the validator
        /// </summary>
        public PptxValidator()
        {
            try
            {
            this.settings = new XmlReaderSettings();
            //this.report = report;

            // resolver
            EmbeddedResourceResolver resolver = new EmbeddedResourceResolver(Assembly.GetEntryAssembly(),
                "CleverAge.OdfConverter.CommandLineTool", ".resources.", true);
            this.settings.XmlResolver = resolver;

            // schemas
            this.settings.Schemas.XmlResolver = resolver;
            
            this.settings.Schemas.Add(OOX_RELATIONSHIP_NS, OOX_RELATIONSHIP_SCHEMA);
            this.settings.Schemas.Add(OOX_CONTENT_TYPE_NS, OOX_CONTENT_TYPE_SCHEMA);
            //this.settings.Schemas.Add(OOX_META_CORE_NS, OOX_META_CORE_SCHEMA);
            this.settings.Schemas.Add(OOX_META_APP_NS, OOX_META_APP_SCHEMA);
                 
            
            this.settings.ValidationType = ValidationType.Schema;
            this.settings.ValidationEventHandler += new ValidationEventHandler(ValidationHandler);
        }
        catch (Exception ex)
        {
            Console.Write(ex.Message);
        }
        }

        /// <summary>
        /// Check the validity of an Office Open XML file.
        /// </summary>
        /// <param name="fileName">The path of the docx file.</param>
        public void validate(String fileName)
        {

            // 0. The file must exist and be a valid Zip archive
            
            try
            {
                reader = ZipFactory.OpenArchive(fileName);
            }
            catch (Exception e)
            {
                throw new PptxValidatorException("Problem opening the pptx file : " + e.Message);
            }

            // 1. [Content_Types].xml must be present and valid
            Stream contentTypes = null;
            try
            {
                contentTypes = reader.GetEntry(OOX_CONTENT_TYPE_FILE);
            }
            catch (Exception)
            {
                throw new PptxValidatorException("The pptx package must have a \"/[Content_Types].xml\" file");
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
                throw new PptxValidatorException("The pptx package must have a \"/_rels/.rels\" file");
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
                throw new PptxValidatorException("openDocument relation not found in \"/_rels/.rels\"");
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
                        throw new PptxValidatorException("The file \"" + target + "\" is described in the \"/_rels/.rels\" file but does not exist in the package.");
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

            if (partRelExists)
            {
                this.validateXml(partRel);

                // 5. For each item in /ppt/_rels/presentation.xml.rels
                partRel = reader.GetEntry(partRelPath);
                r = XmlReader.Create(partRel);
                ValidateRels(partDir, r);

                //retrieve all ids referenced in the document
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
                        throw new PptxValidatorException("Referenced id exist but no part relationship file found");
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
                        throw new PptxValidatorException("One or more relationship id have not been found in the partRelationship file : " + ids[0]);
                    }
                }
            }

            //6. For each item in /ppt/slideMasters/_rels/slideMasters.xml.rels
            docTarget = null;
            partRel = reader.GetEntry(partRelPath);
            r = XmlReader.Create(partRel);
            while (r.Read() && docTarget == null)
            {
                if (r.NodeType == XmlNodeType.Element && r.GetAttribute("Type") == OOX_SIDEMASTER_RELATIONSHIP_TYPE)
                {
                    docTarget = r.GetAttribute("Target");
                    partDir = OOX_SIDEMASTER_RELATIONSHIP_FILE;
                    partRelPath = partDir + "/_rels/" + docTarget.Substring(docTarget.IndexOf("/") + 1) + ".rels";
                    
                    partRel = reader.GetEntry(partRelPath);
                    r = XmlReader.Create(partRel);
                    ValidateRels(partDir, r);
                }
            }
            
            //7. Validation for SlideLayouts if required will be added here.
 
            //8. For each slide in a presentation, validation needs to be added here

            //9. During developement required validations for pptx feature e.g notes master, numbering etc. will be added.

            //10. are all referenced relationships exist in the part relationship file





            
        }

        // validate xml stream
        private void validateXml(Stream xmlStream)
        {
            XmlReader r = XmlReader.Create(xmlStream, this.settings);
            while (r.Read()) ;
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
                    }
                    else if (r.LocalName == "Override" && r.GetAttribute("PartName") == target)
                    {
                        overrided = true;
                        contentType = r.GetAttribute("ContentType");
                    }
                }
            }
            if (contentType == null)
            {
                throw new PptxValidatorException("Content type not found for " + target);
            }
            return contentType;
        }

        public void ValidateRels(string partDir, XmlReader r)
        {
                           
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
                        throw new PptxValidatorException("The file \"" + target + "\" is described in the \"/ppt/_rels/presentation.xml.rels\" file but does not exist in the package.");
                        fileExists = false;
                    }

                    if (fileExists)
                    {
                        
                        // 5.1. A content type can be found in [Content_Types].xml file
                        if (target.IndexOf("../") == -1)
                        {
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
            
           
        }

        public void ValidationHandler(object sender, ValidationEventArgs args)
        {
            throw new PptxValidatorException("XML Schema Validation error : " + args.Message);
        }
    }
}
