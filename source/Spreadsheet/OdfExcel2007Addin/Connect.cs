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

namespace CleverAge.OdfConverter.OdfExcel2007Addin
{
    using System;
    using System.Reflection;
    using Microsoft.Office.Core;
    using System.IO;
    using System.Xml;
    using Extensibility;
    using System.Runtime.InteropServices;
    using System.Collections;
    using MSExcel = Microsoft.Office.Interop.Excel;
    using CleverAge.OdfConverter.OdfConverterLib;
    using CleverAge.OdfConverter.OdfExcelAddinLib;

	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the MyAddin1Setup project, 
	// right click the project in the Solution Explorer, then choose install.
	#endregion
	
	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("ACA16CE6-4077-4DD6-825F-CA913B9649A5"), ProgId("OdfExcel2007Addin.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2,IRibbonExtensibility, IOdfConverter
	{
        private const string ODF_FILE_TYPE = "OdfExcelFileType";
        private const string ALL_FILE_TYPE = "AllFileType";
        private const string IMPORT_ODF_FILE_FILTER = "*.ods";
        private const string IMPORT_ALL_FILE_FILTER = "*.*";
        private const string IMPORT_LABEL = "OdfImportLabel";
        private const string EXPORT_ODF_FILE_FILTER = " (*.ods)|*.ods";
        //private const string EXPORT_ALL_FILE_FILTER = " (*.*)|*.*";
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
            this.addinLib = new CleverAge.OdfConverter.Spreadsheet.Addin();
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
            this.applicationObject = (MSExcel.Application)application;
            
            // set culture to match current application culture or user's choice
            int culture = 0;
            string languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_CURRENT_USER\Software\Clever Age\Odf Add-in for Excel", "Language", null) as string;

            if (languageVal == null)
            {
                languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_LOCAL_MACHINE\Software\Clever Age\Odf Add-in for Excel", "Language", null) as string;
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
        /// Read an ODF file
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public void ImportODF(IRibbonControl control)
        {
            FileDialog fd = this.applicationObject.get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
            // allow multiple file opening
            fd.AllowMultiSelect = true;
            // add filter for ODS files
            fd.Filters.Clear();
            fd.Filters.Add(this.addinLib.GetString(ODF_FILE_TYPE), IMPORT_ODF_FILE_FILTER, Type.Missing);
            // set title
            fd.Title = this.addinLib.GetString(IMPORT_LABEL);
            // display the dialog
            fd.Show();
            
            try {
                // process the chosen documents	
                for (int i = 1; i <= fd.SelectedItems.Count; ++i) {
                    // retrieve file name
                    String odfFile = fd.SelectedItems.Item(i);

                    // create a temporary file
                    object fileName = this.addinLib.GetTempFileName(odfFile, ".xlsx");

                    // this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;
                    OdfToOox(odfFile, (string)fileName, true);
                    // this.applicationObject.System.Cursor =  WdCursorType.wdCursorNormal;

                    try {
                        // open the workbook
                        object readOnly = true;
                        object addToRecentFiles = false;
                        object isVisible = true;
                        object openAndRepair = false;
                        object missing = Missing.Value;

                        // conversion may have been cancelled and file deleted.
                        if (File.Exists((string)fileName)) {
                            // Workaround to excel bug. "Old format or invalid type library. (Exception from HRESULT: 0x80028018 (TYPE_E_INVDATAREAD))" 
                            System.Globalization.CultureInfo ci;
                            ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                            Microsoft.Office.Interop.Excel.Workbook wb =
                                this.applicationObject.Workbooks.Open((string)fileName, missing, readOnly , missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

                            wb.Activate();

                            System.Threading.Thread.CurrentThread.CurrentCulture = ci;
                        }
                    } catch (Exception ex) {
                        // this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorNormal;
                        System.Diagnostics.Debug.WriteLine("*** Exception : " + ex.Message);
                        System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfUnexpectedError"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                        return;
                    }
                }
            } finally {                
            }
        }

        /// <summary>
        /// Save as ODF.
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public void ExportODF(IRibbonControl control)
        {
            // Workaround to excel bug. "Old format or invalid type library. (Exception from HRESULT: 0x80028018 (TYPE_E_INVDATAREAD))" 
            System.Globalization.CultureInfo ci;
            ci = System.Threading.Thread.CurrentThread.CurrentCulture;           
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                MSExcel.Workbook wb = this.applicationObject.ActiveWorkbook;

                // the second test deals with blank workbooks
                // (which are in a 'saved' state and have no extension yet(?))
                if (!wb.Saved || wb.FullName.IndexOf('.') < 0)
                {
                    System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfSaveDocumentBeforeExport"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                }
                else
                {
                    System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
                    // sfd.SupportMultiDottedExtensions = true;
                    sfd.AddExtension = true;
                    sfd.DefaultExt = "ods";
                    sfd.Filter = this.addinLib.GetString(ODF_FILE_TYPE) + EXPORT_ODF_FILE_FILTER;
                                // + this.addinLib.GetString(ALL_FILE_TYPE) + EXPORT_ALL_FILE_FILTER;
                    sfd.InitialDirectory = wb.Path;
                    sfd.OverwritePrompt = true;
                    sfd.Title = this.addinLib.GetString(EXPORT_LABEL);
                    string ext = '.' + sfd.DefaultExt;
                    sfd.FileName = wb.FullName.Substring(0, wb.FullName.LastIndexOf('.')) + ext;

                    // process the chosen workbooks
                    if (System.Windows.Forms.DialogResult.OK == sfd.ShowDialog())
                    {
                        // retrieve file name
                        string odfFile = sfd.FileName;
                        if (!odfFile.EndsWith(ext))
                        { // multi dotted extension support
                            odfFile += ext;
                        }
                        object initialName = wb.FullName;
                        object ooxTempFileName = null;
                        string xlsxFile = (string)initialName;

                        //this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;

                        if (wb.FileFormat != MSExcel.XlFileFormat.xlOpenXMLWorkbook)
                        {
                            // duplicate the file
                            object wbCopyName = Path.GetTempFileName() + Path.GetExtension((string)initialName);
                            File.Copy((string) initialName, (string) wbCopyName);

                            //converting readonly files 
                            FileInfo lFi;

                            lFi = new FileInfo((string)wbCopyName);

                            if (lFi.IsReadOnly)
                            {
                                lFi.IsReadOnly = false;
                            }

                            // open the duplicated file
                            object addToRecentFiles = false;
                            object readOnly = false;
                            object isVisible = false;
                            object missing = Missing.Value;

                            MSExcel.Workbook wbCopy = this.applicationObject.Workbooks.Open((string) wbCopyName, missing, readOnly, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                            
                            // generate xlsx file from the duplicated file (under a temporary file)
                            ooxTempFileName = this.addinLib.GetTempPath((string) initialName, ".xlsx");

                            wbCopy.SaveAs((string) ooxTempFileName, MSExcel.XlFileFormat.xlOpenXMLWorkbook, missing, missing, missing, missing, MSExcel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                           
                            // close and remove the duplicated file
                            wbCopy.Close(false, false, missing);
                            try
                            {
                                File.Delete((string) wbCopyName);
                            }
                            catch (Exception)
                            {
                                // bug #1610099
                                // deletion failed : file currently used by another application.
                                // do nothing
                                // bug #1707349 
                            }
                            xlsxFile = (string) ooxTempFileName;
                        }

                        OoxToOdf(xlsxFile, odfFile, true);

                        if (ooxTempFileName != null && File.Exists((string) ooxTempFileName))
                        {
                            this.addinLib.DeleteTempPath((string)ooxTempFileName);
                        }
                    }
                }
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = ci;                
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

        /// <summary>
        /// Get custom UI
        /// </summary>
        /// <param name="RibbonID">the ribbon identifier</param>
        /// <returns>string ui</returns>
        string IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            using (System.IO.TextReader tr = new System.IO.StreamReader(GetCustomUI()))
            {
                return tr.ReadToEnd();
            }
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

        /// <summary>
        /// ODF Options
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public void ODFOptions(IRibbonControl control)
        {
            using (ConfigForm cfgForm = new ConfigForm())
            {
                cfgForm.ShowDialog();
            }
        }

        private MSExcel.Application applicationObject;
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