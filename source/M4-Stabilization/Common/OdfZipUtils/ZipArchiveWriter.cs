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

        //Added  by sonata
        /// <summary>
        /// Table of Audio files to be added to the package
        /// </summary>
        private Hashtable audioFilesList;

        //Added  by sonata
        /// <summary>
        /// source attribute of the currently processed audio file
        /// </summary>
        private string audioFileSource;
        //Added  by sonata
        /// <summary>
        /// Target attribute of the currently processed audio file
        /// </summary>
        private string audioFileDestination;

        
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

        protected override void Dispose(bool disposing) {
            if (disposing) {
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
                            CopyAudioFiles();
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
                        	if (binaries != null && !binaries.ContainsKey(binarySource))
                        	{
                        		binaries.Add(binarySource, binaryTarget);
                        	}
                        	binarySource = null;
                        	binaryTarget = null;
                        }                       
                        break;     
                    case EXTRACT_ELEMENT:
                        if (audioFileSource != null && audioFileDestination != null)
                        {
                            if (audioFilesList != null && !audioFilesList.ContainsKey(audioFileSource))
                            {
                                audioFilesList.Add(audioFileSource, audioFileDestination);
                            }
                            audioFileSource = null;
                            audioFileDestination = null;
                        }
                        break;     
                }
         
            }
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            Node elt = (Node)elements.Peek();
 			Debug.WriteLine("[WriteStartAttribute] prefix="+prefix+" localName"+localName+" ns="+ns +" element="+elt.Name);
            
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
			Debug.WriteLine("[WriteEndAttribute] element="+elt.Name);
            
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
                                    audioFilesList = new Hashtable();
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
                                        audioFileSource += text;
                                        Debug.WriteLine("copy source=" + audioFileSource);
                                    }
                                    if (attr.Name.Equals("target"))
                                    {
                                        audioFileDestination += text;
                                        Debug.WriteLine("copy target=" + audioFileDestination);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
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
            if (delegateWriter != null) {
                delegateWriter.Close();
                delegateWriter = null;
            }
            if (zipOutputStream != null) {
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
        		CopyBinary(s, (string) binaries[s]);
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
    		    Debug.WriteLine("CopyBinary : "+source+" --> "+target+", bytes copied = "+bytesCopied);
    	    }           

        }

        //Added by Sonata to Copy the audio files From physical directory to Ziparchive and Vice versa
        private void CopyAudioFiles()
        {
            foreach (string strfile in audioFilesList.Keys)
            {
                CopyAudioFile(strfile, (string)audioFilesList[strfile]);
            }
        }

        //Added by Sonata to Copy the audio files From physical directory to Ziparchive and Vice versa
        /// <summary>
        /// Transfer a audio file between ziparchive and physica directory 
        /// The source archive is handled by the resolver, while the destination archive 
        /// corresponds to our zipOutputStream.  
        /// </summary>
        /// <param name="inputFileSource">File path inside the source archive or physica directory</param>
        /// <param name="outputFileTarget">File path inside the destination archive or physica directory</param>
        private void CopyAudioFile(String inputFileSource, String outputFileTarget)
        {
            string inputFilePath = "";
            string outputFilePath = "";
            string targetFolderName = "";
            string targetFileName = "";

            bool ConvertPPTX2ODP = false;
            bool isBatchODP = false;
            bool isBatchPPTX = false;

            // Gets the Target folder name - Copying the .wav files for ODP
            if (outputFileTarget.Contains("|"))
            {
                targetFolderName = outputFileTarget.Substring(0, outputFileTarget.IndexOf('|'));
                targetFileName = outputFileTarget.Substring(outputFileTarget.IndexOf('|') + 1);
            }

            #region Get Input and Output path
            for (int i = 0; i < Environment.GetCommandLineArgs().Length; i++)
            {
                if (Environment.GetCommandLineArgs()[i].Contains("/I"))
                    inputFilePath = Environment.GetCommandLineArgs()[i + 1];
                if (Environment.GetCommandLineArgs()[i].Contains("/O"))
                    outputFilePath = Environment.GetCommandLineArgs()[i + 1];
                if (Environment.GetCommandLineArgs()[i].Contains("/BATCH-ODP"))
                    isBatchODP = true;
                if (Environment.GetCommandLineArgs()[i].Contains("/BATCH-PPTX"))
                    isBatchPPTX = true;
            }
            #endregion

            #region For Addin Conversion
            string strCurrentDirectory = Environment.CurrentDirectory;

            if (inputFilePath == "" && outputFilePath == "")
            {                
                if (inputFileSource.Contains("ppt/media"))
                {
                   ConvertPPTX2ODP = true;
                   inputFilePath = strCurrentDirectory + "\\Dummy.pptx";
                }
                else
                {
                    ConvertPPTX2ODP = false;
                    inputFilePath = strCurrentDirectory + "\\Dummy.odp";
                }
            }
            #endregion

            string audioFilePath = "";
            string directoryPath = "";
            string absolutePath = "";
            string relativePath = "";
            int startIndex = 0;
            int endIndex = 0;
            int replaceIndex = 0;

            #region Batch Processing

            //For Batch ODP
            if (isBatchODP)
            {
                if (outputFilePath != "")
                {
                    outputFilePath = outputFilePath.Replace("\\", "//");
                    audioFilePath = outputFilePath + "//_MediaFilesForOdp_" + targetFolderName;
                }
            }

            //For Batch PPTX
            if (isBatchPPTX)
            {
                //to resolve relative path
                if (inputFileSource.Contains(".."))
                {
                    absolutePath = inputFileSource;
                    startIndex = absolutePath.IndexOf("../");
                    endIndex = absolutePath.LastIndexOf("../");
                    replaceIndex = endIndex / 3;

                    relativePath = inputFileSource.Substring(endIndex + 3, absolutePath.Length - (endIndex + 3));
                    absolutePath = "";
                    string[] pathSplitArray = inputFilePath.Replace("\\", "|").Split('|');

                    for (int cnt = 0; cnt < (pathSplitArray.Length - replaceIndex); cnt++)
                    {
                        absolutePath = absolutePath + pathSplitArray[cnt] + "//";
                    }

                    audioFilePath = absolutePath + relativePath.Replace("/", "//");
                    audioFilePath = audioFilePath.Replace("%20", " ");
                }
                else
                {
                    audioFilePath = inputFileSource.Replace("/", "//").Replace("%20", " ");
                }
            }
            #endregion

            #region File Processing
            else
            {
                if (outputFilePath != "")
                {
                    if (outputFilePath.Contains(".odp"))
                    {
                        directoryPath = Path.GetDirectoryName(outputFilePath).Replace("\\", "//");
                        audioFilePath = directoryPath + "//_MediaFilesForOdp_" + targetFolderName;
                    }
                }
                else
                {
                    if (inputFilePath.Contains(".pptx"))
                    {
                        directoryPath = Path.GetDirectoryName(inputFilePath).Replace("\\", "//");
                        audioFilePath = directoryPath + "//_MediaFilesForOdp_" + targetFolderName;
                        ConvertPPTX2ODP = true;
                    }
                    if (inputFilePath.Contains(".odp"))
                    {
                        if (inputFileSource.Contains(".."))
                        {
                            directoryPath = Path.GetDirectoryName(inputFilePath);
                            absolutePath = inputFileSource;
                            startIndex = absolutePath.IndexOf("../");
                            endIndex = absolutePath.LastIndexOf("../");
                            replaceIndex = endIndex / 3;

                            relativePath = inputFileSource.Substring(endIndex + 3, absolutePath.Length - (endIndex + 3));

                            absolutePath = "";
                            string[] pathSplitArray = directoryPath.Replace("\\", "|").Split('|');
                            for (int cnt = 0; cnt < (pathSplitArray.Length - replaceIndex); cnt++)
                            {
                                absolutePath = absolutePath + pathSplitArray[cnt] + "//";
                            }

                            audioFilePath = absolutePath + relativePath.Replace("/", "//");
                            audioFilePath = audioFilePath.Replace("%20", " ");
                        }
                        else
                        {
                            audioFilePath = inputFileSource.Replace("/", "//").Replace("%20", " ");
                        }
                    }
                }
            }
            #endregion

            #region Add files From External Directory to ZipArchive - Vice versa
            // Condition added by lohith.ar@sonata-software.com to check for audio files
            try
            {

                if (ConvertPPTX2ODP)
                {
                    //Copy ppt/Media files to external directory 
                    Stream inputStream = GetStream(inputFileSource);
                    if (inputStream != null)
                    {
                        Directory.CreateDirectory(audioFilePath);
                        audioFilePath = audioFilePath + "//" + targetFileName;
                        if (!File.Exists(audioFilePath))
                        {
                            FileStream fsFile = new FileStream(audioFilePath, FileMode.Create);
                            StreamCopy(inputStream, fsFile);
                        }
                    }
                }
                else
                {
                    //Copy referd audio fiels from external directory to ppt/media
                    if (zipOutputStream != null)
                    {
                        if (File.Exists(audioFilePath))
                        {
                            zipOutputStream.AddEntry(outputFileTarget);
                            FileStream fsSourceFile = new FileStream(audioFilePath, FileMode.Open, FileAccess.Read);

                            int bytesCopied = StreamCopy(fsSourceFile, zipOutputStream);
                            Debug.WriteLine("CopyBinary : " + inputFileSource + " --> " + outputFileTarget + ", bytes copied = " + bytesCopied);
                        }
                    }
                }
            }
            catch (IOException e)
            {
                ZipException exZip = new ZipException(e.Message);
                throw exZip;
            }
            #endregion

        }
        
        private Stream GetStream(string relativeUri) 
        {
        	Uri absoluteUri = resolver.ResolveUri(null, relativeUri);
            return (Stream)resolver.GetEntity(absoluteUri, null, Type.GetType("System.IO.Stream"));
        }
        
        
        private int StreamCopy(Stream source, Stream destination)
        {
        	if (source != null && destination!= null)
        	{
        		byte [] data = new byte[BUFFER_SIZE];
        		int length = (int) source.Length;
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
       

    }

}
