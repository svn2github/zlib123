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
            this.addinLib = new OdfWordAddinLib();
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
            // set culture to match current application culture
            int officeUICulture = this.applicationObject.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(officeUICulture);
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
            fd.Filters.Add(this.addinLib.GetString("OdfFileType"), "*.odt", Type.Missing);
            fd.Filters.Add(this.addinLib.GetString("AllFileType"), "*.*", Type.Missing);
            // set title
            fd.Title = this.addinLib.GetString("OdfImportLabel");
            // display the dialog
            fd.Show();

            // process the chosen documents	
            for (int i = 1; i <= fd.SelectedItems.Count; ++i)
            {
                // retrieve file name
                String odfFile = fd.SelectedItems.Item(i);

                // create a temporary file
                object fileName = this.addinLib.GetTempFileName(odfFile);

                this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorWait;
                OdfToOox(odfFile, (string)fileName, true);
                this.applicationObject.System.Cursor = MSword.WdCursorType.wdCursorNormal;

                // open the document
                object readOnly = true;
                object addToRecentFiles = false;
                object isVisible = true;
                object missing = Type.Missing;
                Microsoft.Office.Interop.Word.Document doc = this.applicationObject.Documents.Open(ref fileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

                // and activate it
                doc.Activate();

            }
		}

        /// <summary>
        /// Save as ODF.
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
		public void ExportODF(IRibbonControl control)
		{
            
            MSword.Document doc = this.applicationObject.ActiveDocument;

            if (!doc.Saved)
            {
                System.Windows.Forms.MessageBox.Show("Please save your document before exporting to ODF");
            }
            else
            {
                System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
                sfd.AddExtension = true;
                sfd.DefaultExt = "odt";
                sfd.Filter = this.addinLib.GetString("OdfFileType") + " (*.odt)|*.odt|"
                             + this.addinLib.GetString("AllFileType") + " (*.*)|*.*";
                sfd.InitialDirectory = doc.Path;
                sfd.OverwritePrompt = true;
                sfd.Title = this.addinLib.GetString("OdfExportLabel");

                // process the chosen documents	
                if (System.Windows.Forms.DialogResult.OK == sfd.ShowDialog())
                {
                    // retrieve file name
                    string odfFile = sfd.FileName; ;

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
                        tmpFileName = Path.GetTempFileName();
                        object format = MSword.WdSaveFormat.wdFormatXMLDocument;
                        newDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                        // close and remove the duplicated file
                        object saveChanges = false;
                        newDoc.Close(ref saveChanges, ref missing, ref missing);
                        File.Delete((string)newName);

                        docxFile = (string)tmpFileName;
                    }

                    OoxToOdf(docxFile, odfFile, true);

                    if (tmpFileName != null && File.Exists((string)tmpFileName))
                    {
                        File.Delete((string)tmpFileName);
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
        private OdfWordAddinLib addinLib;

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
