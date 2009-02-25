/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ``AS IS�� AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Runtime.InteropServices;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Globalization;
using System.Runtime.CompilerServices;
using OdfConverter.Office;
using OdfConverter.OdfConverterLib;

namespace OdfConverter.Spreadsheet.OdfExcelAddin
{
    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [ComVisible(true)]
    [GuidAttribute("E00C9EBB-F140-4E6F-8C7B-EED19AE33AEA")]
    [ProgId("OdfExcelAddin.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : AbstractOdfAddin
    {
        protected const string ODF_FILE_TYPE_ODS = "OdfFileTypeOds";

        protected const string IMPORT_ODF_FILE_FILTER = "*.ods";
        protected const string EXPORT_ODF_FILE_FILTER = " (*.ods)|*.ods|";

        protected const string HKCU_KEY = @"HKEY_CURRENT_USER\Software\OpenXML-ODF Translator\ODF Add-in for Excel";
        protected const string HKLM_KEY = @"HKEY_LOCAL_MACHINE\SOFTWARE\OpenXML-ODF Translator\ODF Add-in for Excel";

        /// <summary>
        /// Class name for Excel12 documents
        /// </summary>
        private const int Excel12Class = 51;

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            this._addinLib = new OdfAddinLib(this, new CleverAge.OdfConverter.Spreadsheet.Converter());
            this._addinLib.OverrideResourceManager = new System.Resources.ResourceManager("OdfExcelAddin.resources.Labels", Assembly.GetExecutingAssembly());
            // Workaround to excel bug. "Old format or invalid type library. (Exception from HRESULT: 0x80028018 (TYPE_E_INVDATAREAD))" 
            if (!CultureInfo.CurrentCulture.Equals(CultureInfo.CurrentUICulture))
            {
                System.Globalization.CultureInfo ci;
                ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            }
        }


        private LateBindingObject OpenDocument(string fileName, bool confirmConversions, bool readOnly, bool addToRecentFiles, bool isVisible, bool openAndRepair)
        {
            LateBindingObject doc = null;
            switch (this._officeVersion)
            {
                case OfficeVersion.OfficeXP:

                    doc = this._application.Invoke("Workbooks").
                    Invoke("Open", fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing);
                    break;

                default:
                    doc = this._application.Invoke("Workbooks").
                    Invoke("Open", fileName, Type.Missing, readOnly, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    break;

            }
            return doc;
        }

        private LateBindingObject SaveDocumentAs(LateBindingObject doc, string fileName)
        {
            // prevent any popup dialogs which block the application in the background
            bool bDisplayAlerts = this._application.GetBool("DisplayAlerts");
            this._application.SetBool("DisplayAlerts", false);

            switch (this._officeVersion)
            {
                case OfficeVersion.OfficeXP:
                    doc.Invoke("SaveAs", fileName, Excel12Class, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing);
                    break;

                case OfficeVersion.Office2003:
                    doc.Invoke("SaveAs", fileName, Excel12Class, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing);
                    break;

                default:
                    doc.Invoke("SaveAs", fileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing);
                    break;
            }
            this._application.SetBool("DisplayAlerts", bDisplayAlerts);

            return doc;
        }

        protected override void InitializeAddin()
        {
        }

        /// <summary>
        /// Read an ODF file.
        /// </summary>
        public override bool importOdfFile(string odfFile)
        {
            try
            {
                bool isTemplate = Path.GetExtension(odfFile).ToUpper().Equals(".OTS");
                // create a temporary file
                string outputExtension = isTemplate ? ".xltx" : ".xlsx";
                string fileName = this._addinLib.GetTempFileName(odfFile, outputExtension);

                ConversionOptions options = new ConversionOptions();
                options.InputFullName = odfFile;
                options.OutputFullName = fileName;
                options.ConversionDirection = ConversionDirection.OdsToXlsx;
                options.Generator = this.GetGenerator();
                options.DocumentType = isTemplate ? DocumentType.Template : DocumentType.Document;
                options.ShowProgress = true;
                options.ShowUserInterface = true;

                this._addinLib.OdfToOox(odfFile, fileName, options);

                // open the document
                bool confirmConversions = false;
                bool readOnly = true;
                bool addToRecentFiles = false;
                bool isVisible = true;
                bool openAndRepair = false;

                // conversion may have been cancelled and file deleted.
                if (File.Exists((string)fileName))
                {
                    // Workaround to excel bug. "Old format or invalid type library. (Exception from HRESULT: 0x80028018 (TYPE_E_INVDATAREAD))" 
                    System.Globalization.CultureInfo ci;
                    ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                    LateBindingObject doc = OpenDocument(fileName, confirmConversions, readOnly, addToRecentFiles, isVisible, openAndRepair);

                    // and activate it
                    doc.Invoke("Activate");

                    System.Threading.Thread.CurrentThread.CurrentCulture = ci;
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
                System.Windows.Forms.MessageBox.Show(this._addinLib.GetString("OdfUnexpectedError"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                return false;
            }
        }

        /// <summary>
        /// Save as ODF.
        /// </summary>
        public override bool ExportOdf()
        {
            System.Globalization.CultureInfo ci;
            ci = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                LateBindingObject doc = _application.Invoke("ActiveWorkbook");

                // the second test deals with blank documents 
                // (which are in a 'saved' state and have no extension yet(?))
                if (!doc.GetBool("Saved")
                    || doc.GetString("FullName").IndexOf('.') < 0
                    || doc.GetString("FullName").IndexOf("http://") == 0
                    || doc.GetString("FullName").IndexOf("https://") == 0
                    || doc.GetString("FullName").IndexOf("ftp://") == 0
                    )
                {
                    System.Windows.Forms.MessageBox.Show(_addinLib.GetString("OdfSaveDocumentBeforeExport"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                    return false;
                }
                else
                {
                    System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();

                    sfd.AddExtension = true;
                    sfd.DefaultExt = "ods";
                    sfd.Filter = this._addinLib.GetString(this.OdfFileType) + this.ExportOdfFileFilter
                                 + this._addinLib.GetString(ALL_FILE_TYPE) + this.ExportAllFileFilter;
                    sfd.InitialDirectory = doc.GetString("Path");
                    sfd.OverwritePrompt = true;
                    sfd.Title = this._addinLib.GetString(EXPORT_LABEL);
                    string ext = '.' + sfd.DefaultExt;
                    sfd.FileName = doc.GetString("FullName").Substring(0, doc.GetString("FullName").LastIndexOf('.')) + ext;

                    // process the chosen documents	
                    if (System.Windows.Forms.DialogResult.OK == sfd.ShowDialog())
                    {
                        // name of the file to create
                        string odfFileName = sfd.FileName;
                        // multi dotted extensions support
                        if (!odfFileName.EndsWith(ext))
                        {
                            odfFileName += ext;
                        }
                        // name of the document to convert
                        string sourceFileName = doc.GetString("FullName");
                        // name of the temporary Word12 file created if current file is not already a Word12 document
                        string tempXlsxName = null;

                        if (doc.GetInt32("FileFormat") != (int)XlFileFormat.xlOpenXMLWorkbook
                            || doc.GetInt32("FileFormat") != (int)XlFileFormat.xlOpenXMLWorkbookMacroEnabled
                            || doc.GetInt32("FileFormat") != (int)XlFileFormat.xlOpenXMLTemplate
                            || doc.GetInt32("FileFormat") != (int)XlFileFormat.xlOpenXMLTemplateMacroEnabled)
                        {
                            // if file is not currently in Excel12 format
                            // 1. Create a copy
                            // 2. Open it and do a "Save as Excel12" (copy needed not to perturb current openened document
                            // 3. Convert the Excel12 copy to ODF
                            // 4. Remove both temporary created files

                            // duplicate the file to keep current file "as is"
                            string tempCopyName = Path.GetTempFileName() + Path.GetExtension((string)sourceFileName);
                            File.Copy((string)sourceFileName, (string)tempCopyName);

                            //BUG FIX #1743469
                            FileInfo fi = new FileInfo(tempCopyName);
                            if (fi.IsReadOnly)
                            {
                                fi.IsReadOnly = false;
                            }
                            //BUG FIX #1743469

                            // open the duplicated file
                            bool confirmConversions = false;
                            bool readOnly = false;
                            bool addToRecentFiles = false;
                            bool isVisible = false;

                            LateBindingObject newDoc = OpenDocument(tempCopyName, confirmConversions, readOnly, addToRecentFiles, isVisible, false);

                            newDoc.Invoke("Windows").Invoke("Item", 1).SetBool("Visible", false);

                            // generate xlsx file from the duplicated file (under a temporary file)
                            tempXlsxName = this._addinLib.GetTempPath((string)sourceFileName, ".xlsx");

                            SaveDocumentAs(newDoc, tempXlsxName);

                            // close and remove the duplicated file
                            newDoc.Invoke("Close", WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat, Type.Missing);

                            //BUG FIX #1743469
                            try
                            {
                                File.Delete((string)tempCopyName);
                            }
                            catch (Exception ex)
                            {
                                //If delete does not work, don't stop the rest of the process
                                //The tempFile will be deleted by the system
                                System.Diagnostics.Trace.WriteLine(ex.ToString());
                            }
                            //BUG FIX #1743469

                            // Now the file to be converted is
                            sourceFileName = tempXlsxName;
                        }

                        ConversionOptions options = new ConversionOptions();
                        options.InputFullName = sourceFileName;
                        options.InputFullNameOriginal = doc.GetString("FullName");
                        options.OutputFullName = odfFileName;
                        options.ConversionDirection = ConversionDirection.XlsxToOds;
                        options.Generator = this.GetGenerator();
                        options.DocumentType = Path.GetExtension(odfFileName).ToUpper().Equals(".OTS") ? DocumentType.Template : DocumentType.Document;
                        options.ShowProgress = true;
                        options.ShowUserInterface = true;

                        this._addinLib.OoxToOdf(sourceFileName, odfFileName, options);

                        if (tempXlsxName != null && File.Exists((string)tempXlsxName))
                        {
                            this._addinLib.DeleteTempPath((string)tempXlsxName);
                        }
                    }
                    return true;
                }
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            }
        }

        protected override string OdfFileType
        {
            get { return ODF_FILE_TYPE_ODS; }
        }

        protected override string ImportOdfFileFilter
        {
            get { return IMPORT_ODF_FILE_FILTER; }
        }

        protected override string ExportOdfFileFilter
        {
            get { return EXPORT_ODF_FILE_FILTER; }
        }

        public override string RegistryKeyUser
        {
            get { return HKCU_KEY; }
        }

        public override string RegistryKeyLocalMachine
        {
            get { return HKLM_KEY; }
        }

        protected override string ConnectClassName
        {
            get { return "OdfExcelAddin.Connect"; }
        }
    }
}
