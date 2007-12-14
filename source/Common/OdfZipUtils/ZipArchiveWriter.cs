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
using System.Drawing;
using System.Diagnostics;
using System.Reflection;

namespace CleverAge.OdfConverter.OdfZipUtils
{
    /// <summary>
    /// Zip archiving states
    /// </summary>
    internal enum ProcessingState
    {
        /// <summary>
        /// Not archiving
        /// </summary>
        None,
        /// <summary>
        /// Waiting for an entry
        /// </summary>
        EntryWaiting,
        /// <summary>
        /// Processing an entry
        /// </summary>
        EntryStarted
    }

    /// <summary>
    /// An <c>XmlWriter</c> implementation for serializing the xml stream to a zip archive.
    /// All the necessary information for creating the archive and its entries is picked up 
    /// from the incoming xml stream and must conform to the following specification :
    /// 
    /// TODO : XML schema
    /// 
    /// example :
    /// 
    /// <c>&lt;pzip:archive pzip:target="path"&gt;</c>
    /// 	<c>&lt;pzip:entry pzip:target="relativePath"&gt;</c>
    /// 		<c>&lt;-- xml fragment --&lt;</c>
    /// 	<c>&lt;/pzip:entry&gt;</c>
    /// 	<c>&lt;-- other zip entries --&lt;</c>
    /// <c>&lt;/pzip:archive&gt;</c>
    /// 
    /// </summary>
    public class ZipArchiveWriter : XmlWriter
    {
        private const string ZIP_POST_PROCESS_NAMESPACE = "urn:cleverage:xmlns:post-processings:zip";
        private const string PART_ELEMENT = "entry";
        private const string ARCHIVE_ELEMENT = "archive";
        private const string COPY_ELEMENT = "copy";
        //Added by Sonata
        private const string EXTRACT_ELEMENT = "extract";
        private const string IMPORT_ELEMENT = "import";

        /// <summary>
        /// The zip archive
        /// </summary>
        private ZipWriter zipOutputStream;
        private ProcessingState processingState = ProcessingState.None;
        private Stack elements;
        private Stack attributes;
        /// <summary>
        /// A delegate <c>XmlWriter</c> that actually feeds the zip output stream. 
        /// </summary>
        private XmlWriter delegateWriter = null;
        /// <summary>
        /// The delegate settings
        /// </summary>
        private XmlWriterSettings delegateSettings = null;
        private XmlResolver resolver;
        /// <summary>
        /// Source attribute of the currently processed binary file
        /// </summary>
        private string binarySource;
        /// <summary>
        /// Target attribute of the currently processed binary file
        /// </summary>
        private string binaryTarget;
        /// <summary>
        /// Table of binary files to be added to the package
        /// </summary>
        private Hashtable binaries;

        private struct ImageValue
        {
            public int Width;
            public int Height;
            public float VerticalResolution;
            public float HorizontalResolution;
        }

        private System.Collections.Generic.Dictionary<string,ImageValue> ImageValues = new System.Collections.Generic.Dictionary<string,ImageValue>();

        #region New Coding for Import and Extract sound files

        //Extract
        //Added  by sonata
        /// <summary>
        /// extractFileList - list of files to be extracted from pptx zip archive
        /// archiveFileSource - input pptx zip archive source
        /// externalFileDestionation - external path to extraxct the media(.wav) files
        /// </summary>
        private Hashtable extractFileList;
        private string archiveFileSource;
        private string externalFileDestionation;

        //Import
        //Added  by sonata
        /// <summary>
        /// importFileList - list of files to be imported from physical directory to PPTX zip archive
        /// externalFileSource - physical directory path for the file to be copied
        /// archiveFileDestination - pptx zip archive destination        
        /// </summary>
        private Hashtable importFileList;
        private string externalFileSource;
        private string archiveFileDestination;

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        public ZipArchiveWriter(XmlResolver res)
        {
            elements = new Stack();
            attributes = new Stack();

            delegateSettings = new XmlWriterSettings();
            delegateSettings.OmitXmlDeclaration = false;
            delegateSettings.CloseOutput = false;
            delegateSettings.Encoding = Encoding.UTF8;
            delegateSettings.Indent = false;
            // If we use a new delegate per entry in the archive, 
            // XML conformance will be checked at the document level.
            // It is not possible to check XML conformance at the docuement level
            // with a single delegate that writes all the entries.
            // We must then use ConformanceLevel.Fragment and the xml declaration will be missing.
            delegateSettings.ConformanceLevel = ConformanceLevel.Document;

            resolver = res;

            //Debug.Listeners.Add(new TextWriterTraceListener("debug.txt"));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose managed resources
                if (delegateWriter != null)
                    delegateWriter.Close();
                if (zipOutputStream != null)
                    zipOutputStream.Dispose();
            }
            base.Dispose(disposing);

        }
        /// <summary>
        /// Delegates <c>WriteStartElement</c> calls when the element's prefix does not 
        /// match a zip command.  
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="localName"></param>
        /// <param name="ns"></param>
        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            Debug.WriteLine("[startElement] prefix=" + prefix + " localName=" + localName + " ns=" + ns);
            elements.Push(new Node(prefix, localName, ns));

            // not a zip processing instruction
            if (!ZIP_POST_PROCESS_NAMESPACE.Equals(ns))
            {
                if (delegateWriter != null)
                {
                    Debug.WriteLine("{WriteStartElement=" + localName + "} delegate");
                    delegateWriter.WriteStartElement(prefix, localName, ns);
                }
                else
                {
                    Debug.WriteLine("[WriteStartElement=" + localName + " } delegate is null");
                }
            }
        }

        /// <summary>
        /// Delegates <c>WriteEndElement</c> calls when the element's prefix does not 
        /// match a zip command. 
        /// Otherwise, close the archive or flush the delegate writer. 
        /// </summary>
        public override void WriteEndElement()
        {
            Node elt = (Node)elements.Pop();

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE))
            {
                Debug.WriteLine("delegate - </" + elt.Name + ">");
                if (delegateWriter != null)
                {
                    delegateWriter.WriteEndElement();
                }
            }
            else
            {
                switch (elt.Name)
                {
                    case ARCHIVE_ELEMENT:
                        if (zipOutputStream != null)
                        {
                            // Copy binaries before closing the archive
                            CopyBinaries();
                            //Added by Sonata - Copy Audio files before closing archive
                            ExtractFiles();
                            ImportFiles();
                            Debug.WriteLine("[closing archive]");
                            zipOutputStream.Close();
                            zipOutputStream = null;
                        }
                        if (processingState == ProcessingState.EntryWaiting)
                        {
                            processingState = ProcessingState.None;
                        }
                        break;
                    case PART_ELEMENT:
                        if (delegateWriter != null)
                        {
                            Debug.WriteLine("[end part]");
                            delegateWriter.WriteEndDocument();
                            delegateWriter.Flush();
                            delegateWriter.Close();
                            delegateWriter = null;
                        }
                        if (processingState == ProcessingState.EntryStarted)
                        {
                            processingState = ProcessingState.EntryWaiting;
                        }
                        break;
                    case COPY_ELEMENT:
                        if (binarySource != null && binaryTarget != null)
                        {
                            if (binaries != null && !binaries.ContainsKey(binaryTarget))
                            {
                                //Target is the key because there is a case where one file has to be
                                //coppied twice to different locations
                                binaries.Add(binaryTarget, binarySource);
                            }
                            binarySource = null;
                            binaryTarget = null;
                        }
                        break;
                    case EXTRACT_ELEMENT:
                        if (archiveFileSource != null && externalFileDestionation != null)
                        {
                            if (extractFileList != null && !extractFileList.ContainsKey(externalFileDestionation))
                            {
                                extractFileList.Add(externalFileDestionation,archiveFileSource);
                            }
                            archiveFileSource = null;
                            externalFileDestionation = null;
                        }
                        break;
                    case IMPORT_ELEMENT:
                        if (externalFileSource != null && archiveFileDestination != null)
                        {
                            if (importFileList != null && !importFileList.ContainsKey(archiveFileDestination))
                            {
                                importFileList.Add(archiveFileDestination, externalFileSource);
                            }
                            externalFileSource = null;
                            archiveFileDestination = null;
                        }
                        break;

                }

            }
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            Node elt = (Node)elements.Peek();
            Debug.WriteLine("[WriteStartAttribute] prefix=" + prefix + " localName" + localName + " ns=" + ns + " element=" + elt.Name);

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteStartAttribute(prefix, localName, ns);
                }
            }
            else
            {
                // we only store attributes of tied to zip processing instructions
                attributes.Push(new Node(prefix, localName, ns));
            }
        }

        public override void WriteEndAttribute()
        {
            Node elt = (Node)elements.Peek();
            Debug.WriteLine("[WriteEndAttribute] element=" + elt.Name);

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteEndAttribute();
                }
            }
            else
            {
                Node attribute = (Node)attributes.Pop();
            }
        }

        // TODO: throw an exception if "target" attribute not set
        public override void WriteString(string text)
        {
            string strDefaultValue = "0";
            //clam: compute correct cropping size
            try
            {
                if (text.Contains("COMPUTEOOXCROPPING[") && (((CleverAge.OdfConverter.OdfZipUtils.ZipArchiveWriter.Node)((CleverAge.OdfConverter.OdfZipUtils.ZipArchiveWriter)this).elements.Peek()).Name == "srcRect"))
                {
                    char[] sep = { ',', ']', '[' };
                    string[] arrValues = text.Split(sep);
                    string filename = arrValues[2];
                    string strValue = arrValues[1];
                    string strType = arrValues[3];
                    strDefaultValue = arrValues[4];
                    double value = 0;

                    System.Globalization.CultureInfo c = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");

                    if (strValue.Contains("cm")) value = Double.Parse(strValue.Replace("cm", ""));

                    System.Threading.Thread.CurrentThread.CurrentCulture = c;


                    if (value != 0)
                    {
                        ImageValue iv;
                        if (ImageValues.ContainsKey(filename))
                        {
                            iv = ImageValues[filename];
                        }
                        else
                        {
                            Stream sourceStream = GetStream(filename);
                            System.Drawing.Bitmap im = new System.Drawing.Bitmap(sourceStream);

                            iv = new ImageValue();
                            iv.Width = im.Width;
                            iv.Height = im.Height;
                            iv.HorizontalResolution = im.HorizontalResolution;
                            iv.VerticalResolution = im.VerticalResolution;

                            ImageValues.Add(filename, iv);
                        }

                        try
                        {
                            switch (strType)
                            {
                                case "t":
                                    text = Math.Floor(value / 2.54 / iv.Height * iv.VerticalResolution * 1000 * 100).ToString();
                                    break;

                                case "b":
                                    text = Math.Floor(value / 2.54 / iv.Height * iv.VerticalResolution * 1000 * 100).ToString();
                                    break;

                                case "l":
                                    text = Math.Floor(value / 2.54 / iv.Width * iv.HorizontalResolution * 1000 * 100).ToString();
                                    break;

                                case "r":
                                    text = Math.Floor(value / 2.54 / iv.Width * iv.HorizontalResolution * 1000 * 100).ToString();
                                    break;

                                default:
                                    break;

                            }

                        }
                        catch (Exception e)
                        {
                            text = strDefaultValue;
                        }

                    }
                    else
                    {
                        text = "0";
                    }

                }
                else if (text.StartsWith("http://www.dialogika.de/odf-converter/makeabspath#"))
                {
                    text = MakeAbsPath(text);
                }
                else if (text.StartsWith("http://www.dialogika.de/odf-converter/makerelpath#"))
                {
                    text = MakeRelPath(text);
                }
                else if (text.Contains("COMPUTEODFCROPPING[") && (((CleverAge.OdfConverter.OdfZipUtils.ZipArchiveWriter.Node)((CleverAge.OdfConverter.OdfZipUtils.ZipArchiveWriter)this).elements.Peek()).Name == "graphic-properties"))
                {

                    string strDefault = "0cm 0cm 0cm 0cm";

                    System.Globalization.CultureInfo c = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");

                    try
                    {
                        char[] sep = { ',', ']', '[' };
                        string[] arrValues = text.Split(sep);
                        string filename = arrValues[5];
                        strDefault = arrValues[6];
                        
                        double r = Double.Parse(arrValues[1]);
                        double l = Double.Parse(arrValues[2]);
                        double t = Double.Parse(arrValues[3]);
                        double b = Double.Parse(arrValues[4]);
                        double r2, l2, t2, b2;

                        ImageValue iv;
                        if (ImageValues.ContainsKey(filename))
                        {
                            iv = ImageValues[filename];
                        }
                        else
                        {
                            Stream sourceStream = GetStream(filename);
                            System.Drawing.Bitmap im = new System.Drawing.Bitmap(sourceStream);

                            iv = new ImageValue();
                            iv.Width = im.Width;
                            iv.Height = im.Height;
                            iv.HorizontalResolution = im.HorizontalResolution;
                            iv.VerticalResolution = im.VerticalResolution;

                            ImageValues.Add(filename, iv);
                        }

                        r2 = r * 2.54 * iv.Width / iv.HorizontalResolution / 1000 / 100;
                        l2 = l * 2.54 * iv.Width / iv.HorizontalResolution / 1000 / 100;
                        t2 = t * 2.54 * iv.Height / iv.VerticalResolution / 1000 / 100;
                        b2 = b * 2.54 * iv.Height / iv.VerticalResolution / 1000 / 100;

                        text = String.Concat("rect(", t2.ToString(), "cm ", r2.ToString(), "cm ", b2.ToString(), "cm ", l2.ToString(), "cm)");

                    }
                    catch (Exception ex)
                    {
                        text = strDefault;
                    }
                    finally
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = c;
                    }
                    
                }
            }
            catch (Exception ex)
            {
                text = strDefaultValue;
            }

            Node elt = (Node)elements.Peek();

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteString(text);
                }
            }
            else
            {
                // Pick up the target attribute 
                if (attributes.Count > 0)
                {
                    Node attr = (Node)attributes.Peek();
                    if (attr.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE))
                    {
                        switch (elt.Name)
                        {
                            case ARCHIVE_ELEMENT:
                                // Prevent nested archive creation
                                if (processingState == ProcessingState.None && attr.Name.Equals("target"))
                                {
                                    Debug.WriteLine("creating archive : " + text);
                                    zipOutputStream = ZipFactory.CreateArchive(text);
                                    processingState = ProcessingState.EntryWaiting;
                                    binaries = new Hashtable();

                                    //Added by sonata
                                    extractFileList = new Hashtable();
                                    importFileList = new Hashtable();
                                }
                                break;
                            case PART_ELEMENT:
                                // Prevent nested entry creation
                                if (processingState == ProcessingState.EntryWaiting && attr.Name.Equals("target"))
                                {
                                    Debug.WriteLine("creating new part : " + text);
                                    zipOutputStream.AddEntry(text);
                                    delegateWriter = XmlWriter.Create(zipOutputStream, delegateSettings);
                                    processingState = ProcessingState.EntryStarted;
                                    delegateWriter.WriteStartDocument();
                                }
                                break;
                            case COPY_ELEMENT:
                                if (processingState != ProcessingState.None)
                                {
                                    if (attr.Name.Equals("source"))
                                    {
                                        binarySource += text;
                                        Debug.WriteLine("copy source=" + binarySource);
                                    }
                                    if (attr.Name.Equals("target"))
                                    {
                                        binaryTarget += text;
                                        Debug.WriteLine("copy target=" + binaryTarget);
                                    }
                                }
                                break;
                            case EXTRACT_ELEMENT:
                                if (processingState != ProcessingState.None)
                                {
                                    if (attr.Name.Equals("source"))
                                    {
                                        archiveFileSource += text;
                                        Debug.WriteLine("copy source=" + archiveFileSource);
                                    }
                                    if (attr.Name.Equals("target"))
                                    {
                                        externalFileDestionation += text;
                                        Debug.WriteLine("copy target=" + externalFileDestionation);
                                    }
                                }
                                break;
                            case IMPORT_ELEMENT:
                                if (processingState != ProcessingState.None)
                                {
                                    if (attr.Name.Equals("source"))
                                    {
                                        externalFileSource += text;
                                        Debug.WriteLine("copy source=" + externalFileSource);
                                    }
                                    if (attr.Name.Equals("target"))
                                    {
                                        archiveFileDestination += text;
                                        Debug.WriteLine("copy target=" + archiveFileDestination);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }

        private string MakeAbsPath(string text)
        {
            int pos = "http://www.dialogika.de/odf-converter/makeabspath#".Length;
            string relativePath = text.Remove(0, pos);
            string path = "";

            // for Commandline tool
            for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
            {
                if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/I")
                    path = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
            }
            //for addins
            if (path == "")
            {
                path = Environment.CurrentDirectory;
            }


            string absolutePath = Path.GetFullPath(path + relativePath).Replace(" ", "%20");
            return "file:///" + absolutePath;
        }

        private string MakeRelPath(string text)
        {
            int pos = "http://www.dialogika.de/odf-converter/makerelpath#".Length;
            string absolutePath = text.Remove(0, pos);
            absolutePath = absolutePath.Replace("%20", " ");
            string path = "";

            //do not change network paths 
            if (absolutePath.StartsWith("\\\\"))
            {
                return absolutePath;
            }
            else
            {
                // for Commandline tool
                for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
                {
                    if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/I")
                        path = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
                }
                //for addins
                if (path == "")
                {
                    path = Environment.CurrentDirectory;
                }
                string relativePath = MakeWinPathRelative(path, absolutePath);
                return relativePath.Replace("\\", "/");
            }           
        }

        private string MakeWinPathRelative(string absolutePath, string relativeTo)
        {
            string[] absoluteDirectories = absolutePath.Split('\\');
            string[] relativeDirectories = relativeTo.Split('\\');

            //Get the shortest of the two paths
            int length = absoluteDirectories.Length < relativeDirectories.Length ? absoluteDirectories.Length : relativeDirectories.Length;

            //Use to determine where in the loop we exited
            int lastCommonRoot = -1;
            int index;

            //Find common root
            for (index = 0; index < length; index++)
                if (absoluteDirectories[index] == relativeDirectories[index])
                    lastCommonRoot = index;
                else
                    break;

            //If we didn't find a common prefix then throw
            if (lastCommonRoot == -1)
                throw new ArgumentException("Paths do not have a common base");

            //Build up the relative path
            StringBuilder relativePath = new StringBuilder();

            //Add on the ..
            for (index = lastCommonRoot + 1; index < absoluteDirectories.Length; index++)
                if (absoluteDirectories[index].Length > 0)
                    relativePath.Append("..\\");

            //Add on the folders
            for (index = lastCommonRoot + 1; index < relativeDirectories.Length - 1; index++)
                relativePath.Append(relativeDirectories[index] + "\\");
            relativePath.Append(relativeDirectories[relativeDirectories.Length - 1]);

            return "../" + relativePath.ToString();
        }

        public override void WriteFullEndElement()
        {
            this.WriteEndElement();
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
            if (delegateWriter != null)
            {
                delegateWriter.WriteCData(s);
            }
        }

        public override void WriteComment(string s)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteComment(s);
            }
        }

        public override void WriteProcessingInstruction(string name, string text)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteProcessingInstruction(name, text);
            }
        }

        public override void WriteEntityRef(String name)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteEntityRef(name);
            }
        }

        public override void WriteCharEntity(char c)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteCharEntity(c);
            }
        }

        public override void WriteWhitespace(string s)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteWhitespace(s);
            }
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteSurrogateCharEntity(lowChar, highChar);
            }
        }

        public override void WriteChars(char[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteChars(buffer, index, count);
            }
        }

        public override void WriteRaw(char[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteRaw(buffer, index, count);
            }
        }

        public override void WriteRaw(string data)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteRaw(data);
            }
        }

        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteBase64(buffer, index, count);
            }
        }

        public override WriteState WriteState
        {
            get
            {
                if (delegateWriter != null)
                {
                    return delegateWriter.WriteState;
                }
                else
                {
                    return WriteState.Start;
                }
            }
        }

        public override void Close()
        {
            // zipStream and delegate are closed elsewhere.... if everything else is fine
            if (delegateWriter != null)
            {
                delegateWriter.Close();
                delegateWriter = null;
            }
            if (zipOutputStream != null)
            {
                zipOutputStream.Close();
                zipOutputStream = null;
            }

        }

        public override void Flush()
        {
            // nothing to do
        }

        public override string LookupPrefix(String ns)
        {
            if (delegateWriter != null)
            {
                return delegateWriter.LookupPrefix(ns);
            }
            else
            {
                return null;
            }
        }

        private void CopyBinaries()
        {
            foreach (string s in binaries.Keys)
            {
                CopyBinary((string)binaries[s], s);
            }
        }

        private const int BUFFER_SIZE = 4096;

        /// <summary>
        /// Transfer a binary file between two zip archives. 
        /// The source archive is handled by the resolver, while the destination archive 
        /// corresponds to our zipOutputStream.  
        /// </summary>
        /// <param name="source">Relative path inside the source archive</param>
        /// <param name="target">Relative path inside the destination archive</param>
        private void CopyBinary(String source, String target)
        {
            Stream sourceStream = GetStream(source);

            if (sourceStream != null && zipOutputStream != null)
            {
                // New file entry
                zipOutputStream.AddEntry(target);
                int bytesCopied = StreamCopy(sourceStream, zipOutputStream);
                Debug.WriteLine("CopyBinary : " + source + " --> " + target + ", bytes copied = " + bytesCopied);
            }
        }

        #region new code for Extract and Import

        //Added by Sonata
        private void ExtractFiles()
        {
            foreach (string strfile in extractFileList.Keys)
            {
                ExtractFile(strfile, (string)extractFileList[strfile]);
            }
        }

        /// <summary>
        /// Extract files from the input PPTX zip archive to the specified physical directory
        /// The source archive is handled by the resolver, while the destination correspond to physical dir 
        /// </summary>
        /// <param name="source">Relative path inside the source archive</param>
        /// <param name="destination">Destination folder name and file name seperated by '|' character</param>
        private void ExtractFile(string destination, string source)
        {
            // Retrive the Target folder name and Filename
            string targetFolderName = destination.Substring(0, destination.IndexOf('|'));
            string targetFileName = destination.Substring(destination.IndexOf('|') + 1);

            //GetOutputFilepath - gets the path where the output ODP file will be copied
            string strCurrentDirectory = GetOutputFilePath();
            string strDestinationFilePath = strCurrentDirectory.Replace("\\", "//") + "//" + targetFolderName;

            //get the zipArchive stream
            try
            {
                Stream inputStream = GetStream(source);
                if (inputStream != null)
                {
                    Directory.CreateDirectory(strDestinationFilePath);
                    strDestinationFilePath = strDestinationFilePath + "//" + targetFileName;
                    if (!File.Exists(strDestinationFilePath))
                    {
                        FileStream fsFile = new FileStream(strDestinationFilePath, FileMode.Create);
                        StreamCopy(inputStream, fsFile);
                        fsFile.Close();
                    }
                }
            }

            catch (IOException e)
            {
                ZipException exZip = new ZipException(e.Message);
                throw exZip;
            }
        }


        private void ImportFiles()
        {
            foreach (string strfile in importFileList.Keys)
            {
                ImportFile(strfile, (string)importFileList[strfile]);
            }
        }


        /// <summary>
        /// Imoprt files from the specified physical directory to otuput PPTX zip archive
        /// The destination archive is handled by the resolver, while the source corresponds to physical dir 
        /// </summary>
        /// <param name="source">Source file pysical path - Relative/absolute </param>
        /// <param name="destination">Destination path inside zip archive</param>        
        private void ImportFile(string destination, string source)
        {
            string inputFilePath = "";

            //Resolve relative path
            if (source.StartsWith("../"))
            {
                inputFilePath = Path.GetFullPath(Path.Combine(GetInputFilePath(), source.Remove(0, 3))).Replace("/", "//").Replace("%20", " ");
            }
            else
            {
                inputFilePath = source.Replace("/", "//").Replace("%20", " ");
            }

            try
            {
                //Copy referd audio fiels from external directory to ppt/media
                if (zipOutputStream != null)
                {
                    if (File.Exists(inputFilePath))
                    {
                        zipOutputStream.AddEntry(destination);
                        FileStream fsSourceFile = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read);

                        int bytesCopied = StreamCopy(fsSourceFile, zipOutputStream);
                        Debug.WriteLine("CopyBinary : " + inputFilePath + " --> " + destination + ", bytes copied = " + bytesCopied);
                    }
                }
            }

            catch (IOException e)
            {
                ZipException exZip = new ZipException(e.Message);
                throw exZip;
            }
        }

        /// <summary>
        /// Get the physical directory path of the input file (PPTX or ODP) to transformed        /// 
        /// </summary>
        /// <returns>Physical path of the input file </returns>
        private string GetInputFilePath()
        {
            string returnInputFilePath = "";

            // for Commandline tool
            for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
            {
                if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/I")
                    returnInputFilePath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
            }

            //for addinds
            if (returnInputFilePath == "")
            {
                returnInputFilePath = Environment.CurrentDirectory;
            }

            return returnInputFilePath;

        }

        /// <summary>
        /// Get the physical directory path of the output file
        /// </summary>
        /// <returns>Physical path of the output file </returns>
        private string GetOutputFilePath()
        {
            string returnOutputFilePath = "";
            string tempOutputFilePath = "";

            for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
            {
                if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/I")
                    tempOutputFilePath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
                if (Environment.GetCommandLineArgs()[i].ToString().ToUpper() == "/O")
                    returnOutputFilePath = Path.GetDirectoryName(Environment.GetCommandLineArgs()[i + 1]);
            }

            // if /O is not specified /I will remain the pool to copy output file
            if (returnOutputFilePath == "" && tempOutputFilePath != "")
            {
                returnOutputFilePath = tempOutputFilePath;
            }

            //For addins
            else
            {
                returnOutputFilePath = Environment.CurrentDirectory;
            }

            return returnOutputFilePath;
        }

        #endregion

        private Stream GetStream(string source)
        {
            //get stream to embedded resource
            if (source.StartsWith("#CER#") && source.EndsWith("#"))
            {
                int len = source.Length;
                int pos1 = 5;
                int pos2 = source.IndexOf("#", 5, len - 5) + 1;

                Assembly exeAssem = Assembly.GetExecutingAssembly();
                string targetAssemPath = Path.GetDirectoryName(exeAssem.GetModules()[0].FullyQualifiedName);
		if (targetAssemPath != "")
		    targetAssemPath += Path.DirectorySeparatorChar;
		targetAssemPath +=  source.Substring(pos1, pos2 - pos1 -1);
#if STATIC
		if (targetAssemPath.EndsWith(".dll"))
		    targetAssemPath = targetAssemPath.Substring (0, targetAssemPath.Length - 4) + "Static.dll";
#endif
                Assembly targetAssem = System.Reflection.Assembly.LoadFile(targetAssemPath);
                string res = source.Substring(pos2, len - pos2 - 1);

                return targetAssem.GetManifestResourceStream(res);
            }
            //get stream to file in zip
            else
            {
                Uri absoluteUri = resolver.ResolveUri(null, source);
                return (Stream)resolver.GetEntity(absoluteUri, null, Type.GetType("System.IO.Stream"));
            }
        }

        private int StreamCopy(Stream source, Stream destination)
        {
            if (source != null && destination != null)
            {
                byte[] data = new byte[BUFFER_SIZE];
                int length = (int)source.Length;
                int bytesCopied = 0;

                while (length > 0)
                {
                    int read = source.Read(data, 0, System.Math.Min(BUFFER_SIZE, length));
                    bytesCopied += read;

                    if (read < 0)
                    {
                        throw new EndOfStreamException("Unexpected end of stream");
                    }

                    length -= read;
                    destination.Write(data, 0, read);
                }

                source.Close();
                return bytesCopied;
            }

            return -1;
        }

        /// <summary>
        /// Simple representation of elements or attributes nodes
        /// </summary>
        private class Node
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

        //Added by Sonata for Getting Height,Width,Resolution of the image-15/11/2007   
        /// <summary>
        /// Transfer a binary file between two zip archives. 
        /// The source archive is handled by the resolver.
        /// This method returns image properties like height,width, and resolution
        /// </summary>
        /// <param name="source">Relative path inside the source archive</param>
        public string ImageCopyBinary(String source)
        {
            Stream sourceStream = GetStream(source);

            System.Drawing.Image img;
            img = System.Drawing.Image.FromStream(sourceStream);

            return img.Width.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + img.Height.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + img.VerticalResolution.ToString(System.Globalization.CultureInfo.InvariantCulture);

        }
    }

}
