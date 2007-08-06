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

namespace CleverAge.OdfConverter.OdfWord2007Addin
{
    using System;
    using System.Reflection;
    using Microsoft.Office.Core;
    using System.IO;
    using System.Xml;
    using Extensibility;
    using System.Runtime.InteropServices;
    using System.Collections;
    using MSword = Microsoft.Office.Interop.Word;
    using CleverAge.OdfConverter.OdfConverterLib;
    using CleverAge.OdfConverter.OdfWordAddinLib;


	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the ODFWord2007AddinSetup project, 
	// right click the project in the Solution Explorer, then choose install.
	#endregion
	
	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("61835F6F-37E4-423E-9674-B5D070BDA12F")]
    [ProgId("OdfWord2007Addin.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility, IOdfConverter
	{
    	private const string ODF_FILE_TYPE = "OdfFileType";
		private const string ALL_FILE_TYPE = "AllFileType";
		private const string IMPORT_ODF_FILE_FILTER = "*.odt";
		private const string IMPORT_ALL_FILE_FILTER = "*.*";
		private const string IMPORT_LABEL = "OdfImportLabel";
		private const string EXPORT_ODF_FILE_FILTER = " (*.odt)|*.odt|";
		private const string EXPORT_ALL_FILE_FILTER = " (*.*)|*.*";
		private const string EXPORT_LABEL = "OdfExportLabel";
		private string DialogBoxTitle = "ODF Converter";

        [DllImport("kernel32.dll")]
        private static extern void OutputDebugString(string chaine);

        private static void ODS(string chaine)
        {
            OutputDebugString(chaine + "\n");
        }

		/// <summary>
		///		Implements the constructor for the Add-in object.
		///		Place your initialization code within this method.
		/// </summary>
        public Connect()
		{
            this.addinLib = new CleverAge.OdfConverter.Word.Addin();
		}

		/// <summary>
		///      Implements the OnConnection method of the IDTExtensibility2 interface.
		///      Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param term='application'>
		///      Root object of the host application.
		/// </param>
		/// <param term='connectMode'>
		///      Describes how the Add-in is being loaded.
		/// </param>
		/// <param term='addInInst'>
		///      Object representing this Add-in.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            this.applicationObject = (MSword.Application)application;
            // set culture to match current application culture or user's choice
            int culture = 0;
            string languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Word", "Language", null) as string;

            if (languageVal == null)
            {
                languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Clever Age\Odf Add-in for Word", "Language", null) as string;
            }

            if (languageVal != null)
            {
                int.TryParse(languageVal, out culture);
            }

            if (culture == 0)
            {
                culture = this.applicationObject.LanguageSettings
                    .get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            }
            System.Threading.Thread.CurrentThread.CurrentUICulture =
                new System.Globalization.CultureInfo(culture);

		}

		/// <summary>
		///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
		///     Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param term='disconnectMode'>
		///      Describes how the Add-in is being unloaded.
		/// </param>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            this.applicationObject = null;
            this.addinLib = null;
		}

		/// <summary>
		///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		///      Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		/// <summary>
		///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		///      Receives notification that the host application has completed loading.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref System.Array custom)
        {
            // Tell Word that the Normal.dot template should not be saved (unless the user later on makes it dirty)
            this.applicationObject.NormalTemplate.Saved = true;
		}

		/// <summary>
		///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		///      Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref System.Array custom)
		{
		}

        /// <summary>
        /// Get custom UI
        /// </summary>
        /// <param name="RibbonID">the ribbon identifier</param>
        /// <returns>string ui</returns>
		string IRibbonExtensibility.GetCustomUI(string RibbonID)
		{
            using (System.IO.TextReader tr = new System.IO.StreamReader(GetCustomUI())) {
                return tr.ReadToEnd();
            }
		}

		

		
        /// <summary>
        /// Read an ODF file
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public void ImportODF(IRibbonControl control)
        {

            FileDialog fd = this.applicationObject.get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
            // allow multiple file opening
			fd.AllowMultiSelect = true;
            // add filter for ODT files
            fd.Filters.Clear();
            fd.Filters.Add(this.addinLib.GetString(ODF_FILE_TYPE), IMPORT_ODF_FILE_FILTER, Type.Missing);
            fd.Filters.Add(this.addinLib.GetString(ALL_FILE_TYPE), IMPORT_ALL_FILE_FILTER, Type.Missing);
            // set title
            fd.Title = this.addinLib.GetString(IMPORT_LABEL);
            // display the dialog
            fd.Show();

            // process the chosen documents	
            for (int i = 1; i <= fd.SelectedItems.Count; ++i)
            {
                // retrieve file name
                String odfFile = fd.SelectedItems.Item(i);

                // create a temporary file
                object fileName = this.addinLib.GetTempFileName(odfFile, ".docx");

                this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorWait;
                OdfToOox(odfFile, (string)fileName, true);
                this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorNormal;

                try 
                {
                	// open the document
                	object readOnly = true;
                	object addToRecentFiles = false;
                	object isVisible = true;
                	object openAndRepair = false;
                	object missing = Type.Missing;

                    // conversion may have been cancelled and file deleted.
                    if (File.Exists((string) fileName))
                    {
                        Microsoft.Office.Interop.Word.Document doc = this.applicationObject.Documents.Open(ref fileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref openAndRepair, ref missing, ref missing, ref missing);

                        // update document fields
                        doc.Fields.Update();

                        // and activate it
                        doc.Activate();
                    }
                } 
                catch (Exception ex) 
                {
                	this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorNormal;
                    System.Diagnostics.Debug.WriteLine("*** Exception : " + ex.Message);
                   	System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfUnexpectedError"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                  	return;
                }

            }
		}
        
        /// <summary>
        /// Save as ODF.
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
		public void ExportODF(IRibbonControl control)
		{
            
            MSword.Document doc = this.applicationObject.ActiveDocument;

            // the second test deals with blank documents 
            // (which are in a 'saved' state and have no extension yet(?))
            if (!doc.Saved || doc.FullName.IndexOf('.') < 0)
            {
                System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfSaveDocumentBeforeExport"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            }
            else
            {
                System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
                // sfd.SupportMultiDottedExtensions = true;
                sfd.AddExtension = true;
                sfd.DefaultExt = "odt";
                sfd.Filter = this.addinLib.GetString(ODF_FILE_TYPE) + EXPORT_ODF_FILE_FILTER
                             + this.addinLib.GetString(ALL_FILE_TYPE) + EXPORT_ALL_FILE_FILTER;
                sfd.InitialDirectory = doc.Path;
                sfd.OverwritePrompt = true;
                sfd.Title = this.addinLib.GetString(EXPORT_LABEL);
                string ext = '.' + sfd.DefaultExt;
                sfd.FileName = doc.FullName.Substring(0, doc.FullName.LastIndexOf('.')) + ext;

                // process the chosen documents	
                if (System.Windows.Forms.DialogResult.OK == sfd.ShowDialog())
                {
                    // retrieve file name
                    string odfFile = sfd.FileName;
                    if (!odfFile.EndsWith(ext)) 
                    { // multi dotted extension support
                        odfFile += ext;
                    }
                    object initialName = doc.FullName;
                    object tmpFileName = null;
                    string docxFile = (string)initialName;

                    this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorWait;

                    if (doc.SaveFormat != 12)
                    {
                        // duplicate the file
                        object newName = Path.GetTempFileName() + Path.GetExtension((string)initialName);
                        File.Copy((string)initialName, (string)newName);

                        // open the duplicated file
                        object addToRecentFiles = false;
                        object readOnly = false;
                        object isVisible = false;
                        object missing = Type.Missing;

                        MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                        // generate docx file from the duplicated file (under a temporary file)
                        tmpFileName = this.addinLib.GetTempPath((string)initialName, ".docx");
                        object format = MSword.WdSaveFormat.wdFormatXMLDocument;
                        newDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                        // close and remove the duplicated file
                        object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                        object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
                        newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
                        try
                        {
                            File.Delete((string)newName);
                        }
                        catch (IOException)
                        {
                            // bug #1610099
                            // deletion failed : file currently used by another application.
                           // System.Windows.Forms.MessageBox.Show("Failed to delete file " + newName
                           //  , DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                        }
                        docxFile = (string)tmpFileName;
                    }
                        
                    OoxToOdf(docxFile, odfFile, true);
                    
                    if (tmpFileName != null && File.Exists((string)tmpFileName))
                    {
                        this.addinLib.DeleteTempPath((string)tmpFileName);
                    }
                }
            }
		}

        /// <summary>
        /// Get an image
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>IPictureDisp object</returns>
		public stdole.IPictureDisp GetImage(IRibbonControl control)
		{
            return this.addinLib.GetLogo();
		}

        /// <summary>
        /// Get a label
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>label as a string</returns>
		public string getLabel(IRibbonControl control)
		{
            return this.addinLib.GetString(control.Id + "Label");
		}

        /// <summary>
        /// Get description
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>Description as a string</returns>
		public string getDescription(IRibbonControl control)
		{
            return this.addinLib.GetString(control.Id + "Description");
		}

        private Stream GetCustomUI()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("customUI.xml"))
                {
                    return asm.GetManifestResourceStream(name);
                }
            }
            return null;
        }

		private MSword.Application applicationObject;
        private OdfAddinLib addinLib;

        #region IOdfConverter Members

        public void OdfToOox(string inputFile, string outputFile, bool showUserInterface)
        {
            this.addinLib.OdfToOox(inputFile, outputFile, showUserInterface);
        }
        public void OoxToOdf(string inputFile, string outputFile, bool showUserInterface)
        {
            this.addinLib.OoxToOdf(inputFile, outputFile, showUserInterface);
        }

        #endregion
    }
}
