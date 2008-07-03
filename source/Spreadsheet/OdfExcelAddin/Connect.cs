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
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ``AS IS´´ AND ANY
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
        protected const string IMPORT_ODF_FILE_FILTER = "*.ods";
        protected const string EXPORT_ODF_FILE_FILTER = " (*.ods)|*.ods|";

        protected const string HKCU_KEY = @"HKEY_CURRENT_USER\Software\Clever Age\ODF Add-in for Excel";
        protected const string HKLM_KEY = @"HKEY_LOCAL_MACHINE\SOFTWARE\Clever Age\ODF Add-in for Excel";
        
        /// <summary>
        /// Class name for Excel12 documents
        /// </summary>
        private const int Word12Class = 51;

        /// <summary>
        /// Format Id for Excel12 documents in current configuration
        /// </summary>
        private int _word12SaveFormat = -1;

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            this._addinLib = new OdfAddinLib(new CleverAge.OdfConverter.Spreadsheet.Converter());
            this._addinLib.OverrideResourceManager = new System.Resources.ResourceManager("OdfExcelAddin.resources.Labels", Assembly.GetExecutingAssembly());
        }

                
        /// <summary>
        /// Initializes Word12Format field
        /// </summary>
        private int FindWord12SaveFormat()
        {
            int saveFormat = -1;
            try
            {
                if (_officeVersion == OfficeVersion.Office2007)
                {
                    saveFormat = 12;
                }
                else
                {
                    // iterate through file converters to find the correct format
                    LateBindingObject converters = _application.Invoke("FileConverters");

                    for (int i = 1; i <= converters.GetInt32("Count"); i++)
                    {
                        LateBindingObject converter = converters.Invoke("Item", i);
                        string className = converter.GetString("ClassName");
                        if (className.Equals(Word12Class))
                        {
                            // Converter found
                            saveFormat = converter.GetInt32("SaveFormat");
                            break;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            return saveFormat;
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
            //bool addToRecentFiles = false;
            switch (this._officeVersion)
            {
                case OfficeVersion.OfficeXP:
                    doc.Invoke("SaveAs",
                    fileName, Word12Class, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);
                    break;
                
                
                case OfficeVersion.Office2003:
                    doc.Invoke("SaveAs",
                        fileName, Word12Class , Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, 
                        Type.Missing, Type.Missing, Type.Missing);                    
                    break;
                default:
                    doc.Invoke("SaveAs",
                    fileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);
                    break; 
               
            }
            return doc;
        }
        
        protected override void InitializeAddin()
        {
            this._word12SaveFormat = FindWord12SaveFormat();
            
            //// Tell Word that the Normal.dot template should not be saved (unless the user later on makes it dirty)
            //this._application.Invoke("NormalTemplate").SetBool("Saved", true);
        }

        /// <summary>
        /// Read an ODF file.
        /// </summary>
        protected override void importOdf()
        {
            foreach (string odfFile in getOpenFileNames())
        	{
                // create a temporary file
                string fileName = this._addinLib.GetTempFileName(odfFile, ".xlsx");

                this._addinLib.OdfToOox(odfFile, fileName, true);

                try
                {
                    // open the document
                    bool confirmConversions = false;
                    bool readOnly = true;
                    bool addToRecentFiles = false;
                    bool isVisible = true;
                    bool openAndRepair = false;
                    
                    // conversion may have been cancelled and file deleted.
                    if (File.Exists((string)fileName))
                    {
                        LateBindingObject doc = OpenDocument(fileName, confirmConversions, readOnly, addToRecentFiles, isVisible, openAndRepair);
                        
                        // and activate it
                        doc.Invoke("Activate");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.ToString());
                    System.Windows.Forms.MessageBox.Show(this._addinLib.GetString("OdfUnexpectedError"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                    return;
                }
            }
        }

        /// <summary>
        /// Save as ODF.
        /// </summary>
        protected override void exportOdf()
        {

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
            }
            else
            {
                System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();

                sfd.AddExtension = true;
                sfd.DefaultExt = "ods";
                sfd.Filter = this._addinLib.GetString(ODF_FILE_TYPE) + this.ExportOdfFileFilter
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
                    string tempDocxName = null;
                  
                    if (!Path.GetExtension(doc.GetString("FullName")).Equals(".xlsx"))
                    {

                        // duplicate the file
                        string tempCopyName = Path.GetTempFileName() + Path.GetExtension((string)sourceFileName);
                        File.Copy((string)sourceFileName, (string)tempCopyName);

                        // open the duplicated file
                        bool confirmConversions = false;
                        bool readOnly = false;
                        bool addToRecentFiles = false;
                        bool isVisible = false;

                        //Converting readonly files 
                        FileInfo lFi;

                        lFi = new FileInfo((string)tempCopyName);

                        if (lFi.IsReadOnly)
                        {
                            lFi.IsReadOnly = false;
                        }

                        LateBindingObject newDoc = OpenDocument(tempCopyName, confirmConversions, readOnly, addToRecentFiles, isVisible, false);
                        
                        tempDocxName = this._addinLib.GetTempPath((string)sourceFileName, ".xlsx");
                        
                        SaveDocumentAs(newDoc, tempDocxName);
                        
                        sourceFileName = tempDocxName;

                    }

                    this._addinLib.OoxToOdf(sourceFileName, odfFileName, true);

                    if (tempDocxName != null && File.Exists((string)tempDocxName))
                    {
                        this._addinLib.DeleteTempPath((string)tempDocxName);
                    }
                }
            }
        }

        protected override string ImportOdfFileFilter
        {
            get { return IMPORT_ODF_FILE_FILTER; }
        }

        protected override string ExportOdfFileFilter
        {
            get { return EXPORT_ODF_FILE_FILTER; }
        }

        protected override string RegistryKeyUser
        {
            get { return HKCU_KEY; }
        }

        protected override string RegistryKeyLocalMachine
        {
            get { return HKLM_KEY; }
        }

        protected override string ConnectClassName
        {
            get { return "OdfExcelAddin.Connect"; }
        }
    }
}
