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

        protected const string HKCU_KEY = @"HKEY_CURRENT_USER\Software\OpenXML-ODF Translator\ODF Add-in for PowerPoint";
        protected const string HKLM_KEY = @"HKEY_LOCAL_MACHINE\SOFTWARE\OpenXML-ODF Translator\ODF Add-in for PowerPoint";

        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
            this._addinLib = new OdfAddinLib(this, new Sonata.OdfConverter.Presentation.Converter());
            this._addinLib.OverrideResourceManager = new System.Resources.ResourceManager("OdfPowerPointAddin.resources.Labels", Assembly.GetExecutingAssembly());
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
                    Invoke("Open", fileName, Type.Missing, Type.Missing, Type.Missing);
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
                    Invoke("Open", fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                    break;
            }
            return doc;
        }

        private LateBindingObject SaveDocumentAs(LateBindingObject doc, string fileName)
        {
            switch (this._officeVersion)
            {
                case OfficeVersion.Office2007:
                    doc.Invoke("SaveAs", fileName, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                    break;

                default:
                    doc.Invoke("SaveAs", fileName, PpSaveAsFileType.ppSaveAsDefault, Type.Missing);
                    break;
            }
            return doc;
        }

        private LateBindingObject SaveCopyAs(LateBindingObject doc, string fileName)
        {
            switch (this._officeVersion)
            {
                default:
                    doc.Invoke("SaveCopyAs", fileName, PpSaveAsFileType.ppSaveAsTemplate, MsoTriState.msoTrue);
                    break;
            }
            return doc;
        }

        protected override void InitializeAddin()
        {
        }

        /// <summary>
        /// Read an ODF file.
        /// </summary>
        public override void importOdf()
        {
            foreach (string odfFile in getOpenFileNames())
            {
                // create a temporary file
                string fileName = this._addinLib.GetTempFileName(odfFile, ".pptx");

                ConversionOptions options = new ConversionOptions();
                options.InputFullName = odfFile;
                options.OutputFullName = fileName;
                options.ConversionDirection = ConversionDirection.OdpToPptx;
                options.Generator = this.GetGenerator();
                options.DocumentType = Path.GetExtension(odfFile).ToUpper().Equals(".OTP") ? DocumentType.Template : DocumentType.Document;
                options.ShowProgress = true;
                options.ShowUserInterface = true;
                
                this._addinLib.OdfToOox(odfFile, fileName, options);

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
                    System.Diagnostics.Trace.WriteLine(ex.ToString());
                    System.Windows.Forms.MessageBox.Show(this._addinLib.GetString("OdfUnexpectedError"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                    return;
                }
            }
        }

        /// <summary>
        /// Save as ODF.
        /// </summary>
        public override void exportOdf()
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

                    object tmpFileName = null;

                    if (!Path.GetExtension(doc.GetString("FullName")).Equals(".pptx"))
                    {
                        // open the duplicated file
                        object confirmConversions = false;
                        object readOnly = false;
                        object addToRecentFiles = false;
                        object isVisible = false;
                        LateBindingObject newDoc; 

                        switch (this._officeVersion)
                        {
                            case OfficeVersion.OfficeXP:
                                string path = Microsoft.Win32.Registry.GetValue(@"HKEY_CLASSES_ROOT\Powerpoint.Show.12\Shell\Save As\Command", null, null) as string;
                                if (string.IsNullOrEmpty(path))
                                {
                                    System.Windows.Forms.MessageBox.Show(_addinLib.GetString("OdfConverterNotInstalled"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
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
                                break;
                            case OfficeVersion.Office2003:
                                // setting defaultformat value in registry
                                int oldFormat = SetPPTXasDefault();

				                //Code added by Achougle@Xandros on 2 Jan 09
                                tmpFileName = this._addinLib.GetTempPath((string)sourceFileName, ".ppt");
                                //tmpFileName = this._addinLib.GetTempFileName((string)sourceFileName, ".ppt");
                                //Code added by Achougle@Xandros on 2 Jan 09

                                SaveCopyAs(doc, (string)tmpFileName);

                                newDoc = OpenPresentation((string)tmpFileName, confirmConversions, readOnly, addToRecentFiles);

                                SaveDocumentAs(newDoc, (string)tmpFileName);

                                sourceFileName = (string)tmpFileName;

                                // removing defaultformat value in registry
                                RestoreDefault(oldFormat);
                                break;
                            case OfficeVersion.Office2007:
                                // duplicate the file
                                object newName = Path.GetTempFileName() + Path.GetExtension((string)sourceFileName);
                                File.Copy((string)sourceFileName, (string)newName);

                                newDoc = OpenPresentation((string)newName, confirmConversions, readOnly, addToRecentFiles);

                                // generate openxml file from the duplicated file (under a temporary file)
                                tmpFileName = this._addinLib.GetTempPath((string)sourceFileName, ".pptx");

                                SaveDocumentAs(newDoc, (string)tmpFileName);

                                sourceFileName = (string)tmpFileName;
                                break;
                        }

                    }

                    ConversionOptions options = new ConversionOptions();
                    options.InputFullName = sourceFileName;
                    options.InputFullNameOriginal = doc.GetString("FullName");
                    options.OutputFullName = odfFileName;
                    options.ConversionDirection = ConversionDirection.PptxToOdp;
                    options.Generator = this.GetGenerator();
                    options.DocumentType = Path.GetExtension(sourceFileName).ToLower().Equals(".potx")
                        || Path.GetExtension(sourceFileName).ToLower().Equals(".potm") ? DocumentType.Template : DocumentType.Document;
                    options.ShowProgress = true;
                    options.ShowUserInterface = true;
                    this._addinLib.OoxToOdf(sourceFileName, odfFileName, options);
                    
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
