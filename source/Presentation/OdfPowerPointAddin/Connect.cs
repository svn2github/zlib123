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
using Microsoft.Win32;
using System.Diagnostics;

namespace OdfConverter.Presentation.OdfPowerPointAddin
{
    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [ComVisible(true)]
    [GuidAttribute("7f459b4c-65f0-4d44-bb27-66c5fd3ca151")]
    [ProgId("OdfPowerPointAddin.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : AbstractOdfAddin
    {
        protected const string ODF_FILE_TYPE_ODP = "OdfFileTypeOdp";

        protected const string IMPORT_ODF_FILE_FILTER = "*.odp";
        protected const string EXPORT_ODF_FILE_FILTER = " (*.odp)|*.odp|";

        protected const string HKCU_KEY = @"HKEY_CURRENT_USER\Software\OpenXML-ODF Translator\ODF Add-in for Presentation";
        protected const string HKLM_KEY = @"HKEY_LOCAL_MACHINE\SOFTWARE\OpenXML-ODF Translator\ODF Add-in for Presentation";
        
        /// <summary>
        /// Class name for PowerPoint12 documents
        /// </summary>
        private const int PowerPoint12Class = 33;

        /// <summary>
        /// Format Id for PowerPoint12 documents in current configuration
        /// </summary>
        private int _word12SaveFormat = -1;

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()

        {
            this._addinLib = new OdfAddinLib(this, new Sonata.OdfConverter.Presentation.Converter());
            this._addinLib.OverrideResourceManager = new System.Resources.ResourceManager("OdfPowerPointAddin.resources.Labels", Assembly.GetExecutingAssembly());
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
                        if (className.Equals(PowerPoint12Class))
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
                    doc = this._application.Invoke("Presentations").
                    Invoke("Open", fileName, Type.Missing, Type.Missing, Type.Missing);
                    break;

                default:
                    doc = this._application.Invoke("Presentations").
                    Invoke("Open", fileName,Type.Missing,Type.Missing,Type.Missing);
                    break;
            }
            return doc;
        }

        private LateBindingObject OpenPresentation(string fileName, object confirmConversions, object readOnly, object addToRecentFiles)
        {
            LateBindingObject doc = null;
            switch (this._officeVersion)
            {
                case OfficeVersion.OfficeXP:
                    doc = this._application.Invoke("Presentations").
                    Invoke("Open", fileName, Type.Missing, Type.Missing, Type.Missing);
                    break;

                default:
                    doc = this._application.Invoke("Presentations").
                    Invoke("Open", fileName,MsoTriState.msoFalse,MsoTriState.msoTrue,MsoTriState.msoFalse);
                    break;
            }
            return doc;
        }

        private LateBindingObject SaveDocumentAs(LateBindingObject doc, string fileName)
        {
            switch (this._officeVersion)
            {
                case OfficeVersion.Office2007:
                    doc.Invoke("SaveAs",
                         fileName, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                    break;

                default:
                    doc.Invoke("SaveAs",
                        fileName, PpSaveAsFileType.ppSaveAsDefault, Type.Missing);
                    break;
            }
            return doc;
        }

        private LateBindingObject SaveCopyAs(LateBindingObject doc, string fileName)
        {
            switch (this._officeVersion)
            {
                default:
                    doc.Invoke("SaveCopyAs",
                        fileName, PpSaveAsFileType.ppSaveAsTemplate,MsoTriState.msoTrue);
                    break;
            }
            return doc;
        }
        
        protected override void InitializeAddin()
        {
            this._word12SaveFormat = FindWord12SaveFormat();                        
        }

        /// <summary>
        /// Read an ODF file.
        /// </summary>
        protected override void importOdf()
        {
            foreach (string odfFile in getOpenFileNames())
        	{
                // create a temporary file
                string fileName = this._addinLib.GetTempFileName(odfFile, ".pptx");

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
            LateBindingObject doc = _application.Invoke("ActivePresentation");

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
                sfd.DefaultExt = "odp";
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
                        string tempDocxName = null;
                        object tmpFileName = null;

                        if (!Path.GetExtension(doc.GetString("FullName")).Equals(".pptx"))
                        {
                            string officeVersion = _officeVersion.ToString();


                            // open the duplicated file
                            object confirmConversions = false;
                            object readOnly = false;
                            object addToRecentFiles = false;
                            object isVisible = false;
                            
                            if (officeVersion == "OfficeXP")
                            {
                                string path = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\Powerpoint.Show.12\Shell\Save As\Command", null, null) as string;
                                if (string.IsNullOrEmpty(path))
                                {
                                    System.Windows.Forms.MessageBox.Show(_addinLib.GetString("OdfConverterNotInstalled"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                                    //System.Windows.Forms.MessageBox.Show("Microsoft Office 2007 Compatibility Pack not installed or obsolete");
                                }
                                else
                                {
                                    tmpFileName = this._addinLib.GetTempFileName((string)sourceFileName, ".pptx");
                                    path = path.ToLower();
                                    int start = path.IndexOf("moc.exe");
                                    path = path.Substring(0, start) + "ppcnvcom.exe";
                                    string parms = "-oice \"" + doc.GetString("FullName") + "\" \"" + tmpFileName + "\"";
                                    Process p = Process.Start(path, parms);
                                    p.WaitForExit();
                                    sourceFileName = (string)tmpFileName;
                                }
                            }
                            if (officeVersion == "Office2003")
                            {
                                
                                // setting defaultformat value in registry
                                int oldFormat = SetPPTXasDefault();

                                tmpFileName = this._addinLib.GetTempFileName((string)sourceFileName, ".ppt");

                                SaveCopyAs(doc, (string)tmpFileName);

                                LateBindingObject newDoc1 = OpenPresentation((string)tmpFileName, confirmConversions, readOnly, addToRecentFiles);

                                SaveDocumentAs(newDoc1, (string)tmpFileName);

                                sourceFileName = (string)tmpFileName;

                                // removing defaultformat value in registry
                                RestoreDefault(oldFormat);
                            }

                            if (officeVersion == "Office2007")
                            {

                                // duplicate the file
                                object newName = Path.GetTempFileName() + Path.GetExtension((string)sourceFileName);
                                File.Copy((string)sourceFileName, (string)newName);

                                LateBindingObject newDoc1 = OpenPresentation((string)newName, confirmConversions, readOnly, addToRecentFiles);

                                // generate openxml file from the duplicated file (under a temporary file)
                                tmpFileName = this._addinLib.GetTempPath((string)sourceFileName, ".pptx");

                                SaveDocumentAs(newDoc1, (string)tmpFileName);

                                sourceFileName = (string)tmpFileName;

                            }

                        }

                        this._addinLib.OoxToOdf(sourceFileName, odfFileName, true);
                       

                        if (tmpFileName != null && File.Exists((string)tmpFileName))
                        {
                            this._addinLib.DeleteTempPath((string)tmpFileName);
                        }
                }
                
            }
        }

        protected override string OdfFileType
        {
            get { return ODF_FILE_TYPE_ODP; }
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
            get { return "OdfPowerPointAddin.Connect"; }
        }

        int SetPPTXasDefault()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(OptionKey, true))
            {
                int ret = (int)key.GetValue(OptionValue, -1);
                key.SetValue(OptionValue, 0x21, RegistryValueKind.DWord);
                return ret;
            }
        }

        void RestoreDefault(int oldFormat)
        {
            // Delete the value
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(OptionKey, true))
            {
                if (oldFormat == -1)
                {
                    key.DeleteValue(OptionValue);
                }
                else
                {
                    key.SetValue(OptionValue, oldFormat);
                }
            }
        }

        private const string OptionKey = @"Software\Microsoft\Office\11.0\PowerPoint\Options";
        private const string OptionValue = "DefaultFormat";
    }
}
