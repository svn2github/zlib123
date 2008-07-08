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
using System.Collections.Generic;
using System.Text;
using OdfConverter.Office;
using CleverAge.OdfConverter.OdfConverterLib;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;

namespace OdfConverter.OdfConverterLib
{
    /// <summary>
    /// An abstract base class for the Office add-ins
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    public abstract class AbstractOdfAddin : Object, OdfConverter.Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        /// <summary>
        /// The Office version number: 
        ///  9 = Office 2000
        /// 10 = Office XP
        /// 11 = Office 2003
        /// 12 = Office 2007
        /// </summary>
        protected enum OfficeVersion
        {
            Office2000 =  9,
            OfficeXP   = 10,
            Office2003 = 11,
            Office2007 = 12
        }
        
        protected const string ODF_FILE_TYPE = "OdfFileType";
        protected const string ALL_FILE_TYPE = "AllFileType";
        protected const string IMPORT_ALL_FILE_FILTER = "*.*";
        protected const string IMPORT_LABEL = "OdfImportLabel";
        protected const string EXPORT_ALL_FILE_FILTER = " (*.*)|*.*";
        protected const string EXPORT_LABEL = "OdfExportLabel";

        protected string DialogBoxTitle = "ODF Converter";

        /// <summary>
        /// The Office version number: 
        ///  9 = Office 2000
        /// 10 = Office XP
        /// 11 = Office 2003
        /// 12 = Office 2007
        /// </summary>
        protected OfficeVersion _officeVersion = OfficeVersion.Office2007;

        protected LateBindingObject _application;
        protected OdfAddinLib _addinLib;
        protected LateBindingObject /*CommandBarButton*/ _importButton;
        protected LateBindingObject /*CommandBarButton*/ _exportButton;
        protected LateBindingObject /*CommandBarButton*/ _optionsButton;


        protected virtual string[] getOpenFileNames()
        {
            string[] fileNames = null;

            if (_officeVersion > OfficeVersion.Office2000)
            {
                LateBindingObject fileDialog = _application.Invoke("FileDialog", MsoFileDialogType.msoFileDialogFilePicker);

                // allow multiple file opening
                fileDialog.SetBool("AllowMultiSelect", true);

                // add filter for ODT files
                fileDialog.Invoke("Filters").Invoke("Clear");
                fileDialog.Invoke("Filters").Invoke("Add", this._addinLib.GetString(ODF_FILE_TYPE), this.ImportOdfFileFilter, Type.Missing);
                fileDialog.Invoke("Filters").Invoke("Add", this._addinLib.GetString(ALL_FILE_TYPE), this.ImportAllFileFilter, Type.Missing);
                // set title
                fileDialog.SetString("Title", this._addinLib.GetString(IMPORT_LABEL));
                // display the dialog
                fileDialog.Invoke("Show");

                if (fileDialog.Invoke("SelectedItems").GetInt32("Count") > 0)
                {
                    fileNames = new string[fileDialog.Invoke("SelectedItems").GetInt32("Count")];
                    // process the chosen documents	
                    for (int i = 0; i < fileDialog.Invoke("SelectedItems").GetInt32("Count"); i++)
                    {
                        // retrieve file name
                        fileNames[i] = fileDialog.Invoke("SelectedItems").Invoke("Item", i + 1).ToString();
                    }
                }
            }
            else
            {
                // Office 2000 does not provide its own file open dialog 
                // using windows forms instead
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();

                ofd.CheckPathExists = true;
                ofd.CheckFileExists = true;
                ofd.Multiselect = true;
                ofd.SupportMultiDottedExtensions = true;
                ofd.DefaultExt = "odt";
                ofd.Filter = this._addinLib.GetString(ODF_FILE_TYPE) + this.ExportOdfFileFilter
                             + this._addinLib.GetString(ALL_FILE_TYPE) + this.ExportAllFileFilter;

                ofd.Title = this._addinLib.GetString(IMPORT_LABEL);

                // process the chosen documents	
                if (System.Windows.Forms.DialogResult.OK == ofd.ShowDialog())
                {
                    fileNames = ofd.FileNames;
                }
            }

            return fileNames;
        }

        public virtual void SetUICulture()
        {
            // set culture to match current application culture or user's choice
            int culture = 0;
            string languageVal = Microsoft.Win32.Registry.GetValue(this.RegistryKeyUser, "Language", null) as string;

            if (languageVal == null)
            {
                languageVal = Microsoft.Win32.Registry.GetValue(this.RegistryKeyLocalMachine, "Language", null) as string;
            }

            if (languageVal != null)
            {
                int.TryParse(languageVal, out culture);
            }
            if (culture == 0 && _application != null)
            {
                culture = _application.Invoke("LanguageSettings")
                    .Invoke("LanguageID", MsoAppLanguageID.msoLanguageIDUI).ToInt32();
            }

            if (culture != 0)
            {
                System.Threading.Thread.CurrentThread.CurrentUICulture =
                    new System.Globalization.CultureInfo(culture);
            }
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
        public virtual void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            this._application = new LateBindingObject(application);

            // read Office version info
            int version;
            int.TryParse(_application.GetString("Version"),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out version);
            _officeVersion = (OfficeVersion)version;

            Application.EnableVisualStyles();

            this.SetUICulture();

            this.DialogBoxTitle = _addinLib.GetString("OdfConverterTitle");

            if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
            {
                OnStartupComplete(ref custom);
            }
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
        public virtual void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
            {
                OnBeginShutdown(ref custom);
            }
            this._application.Dispose();
            this._addinLib = null;
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public virtual void OnAddInsUpdate(ref System.Array custom)
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
        public virtual void OnStartupComplete(ref System.Array custom)
        {
            if (_officeVersion < OfficeVersion.Office2007)
            {
                // Add menu item
                // first retrieve "File" menu
                LateBindingObject commandBar = _application.Invoke("CommandBars", "File");

                // Add import button
                try
                {
                    // if item already exists, use it (should never happen)
                    _importButton = commandBar.Invoke("Controls", this._addinLib.GetString("OdfImportLabel"));
                }
                catch (Exception)
                {
                    // otherwise, create a new one
                    _importButton = commandBar.Invoke("Controls")
                        .Invoke("Add", MsoControlType.msoControlButton, Type.Missing, Type.Missing, 3, true);
                }
                // set item's label
                _importButton.SetString("Caption", this._addinLib.GetString("OdfImportLabel"));
                _importButton.SetString("Tag", this._addinLib.GetString("OdfImportLabel"));

                // set action
                _importButton.SetString("OnAction", "!<" + this.ConnectClassName + ">");
                _importButton.SetBool("Visible", true);

                _importButton.AddClickEventHandler(new CommandBarButtonEvents_ClickEventHandler(this.importButton_Click));

                // Add export button
                try
                {
                    // if item already exists, use it (should never happen)
                    _exportButton = commandBar.Invoke("Controls", this._addinLib.GetString("OdfExportLabel"));
                }
                catch (Exception)
                {
                    // otherwise, create a new one
                    _exportButton = commandBar.Invoke("Controls")
                        .Invoke("Add", MsoControlType.msoControlButton, Type.Missing, Type.Missing, 4, true);
                }
                // set item's label
                _exportButton.SetString("Caption", this._addinLib.GetString("OdfExportLabel"));
                _exportButton.SetString("Tag", this._addinLib.GetString("OdfExportLabel"));
                // set action
                _exportButton.SetString("OnAction", "!<" + this.ConnectClassName + ">");
                _exportButton.SetBool("Visible", true);
                _exportButton.SetBool("Enabled", true);
                _exportButton.AddClickEventHandler(new CommandBarButtonEvents_ClickEventHandler(this.exportButton_Click));

                // Add options button
                try
                {
                    // if item already exists, use it (should never happen)
                    _optionsButton = commandBar.Invoke("Controls", this._addinLib.GetString("OdfOptionsLabel"));
                }
                catch (Exception)
                {
                    // otherwise, create a new one
                    _optionsButton = commandBar.Invoke("Controls")
                        .Invoke("Add", MsoControlType.msoControlButton, Type.Missing, Type.Missing, 5, true);
                }
                // set item's label
                _optionsButton.SetString("Caption", this._addinLib.GetString("OdfOptionsLabel"));
                _optionsButton.SetString("Tag", this._addinLib.GetString("OdfOptionsLabel"));
                // set action
                _optionsButton.SetString("OnAction", "!<" + this.ConnectClassName + ">");
                _optionsButton.SetBool("Visible", true);
                _optionsButton.SetBool("Enabled", true);
                _optionsButton.AddClickEventHandler(new CommandBarButtonEvents_ClickEventHandler(this.odfOptionsButton_Click));
            }

            this.InitializeAddin();
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public virtual void OnBeginShutdown(ref System.Array custom)
        {
            if (_officeVersion < OfficeVersion.Office2007)
            {
                // remove menu items (use FindControl instead of referenced object
                // in order to actually remove the items - this is a workaround)
                LateBindingObject button = _application.Invoke("CommandBars").Invoke("FindControl", Type.Missing, Type.Missing, this._addinLib.GetString("OdfImportLabel"), Type.Missing);
                button.Invoke("Delete", Type.Missing);

                button = _application.Invoke("CommandBars").Invoke("FindControl", Type.Missing, Type.Missing, this._addinLib.GetString("OdfExportLabel"), Type.Missing);
                button.Invoke("Delete", Type.Missing);

                button = _application.Invoke("CommandBars").Invoke("FindControl", Type.Missing, Type.Missing, this._addinLib.GetString("OdfOptionsLabel"), Type.Missing);
                button.Invoke("Delete", Type.Missing);
            }
        }

        /// <summary>
        /// Event handler for Office 2007
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public virtual void ImportOdf12(IRibbonControl control)
        {
            try
            {
                importOdf();
            }
            catch
            {
            }
        }

        /// <summary>
        /// Event handler for Office versions before 2007
        /// </summary>
        protected virtual void importButton_Click(object ctrl, object CancelDefault)
        {
            importOdf();
        }

        /// <summary>
        /// Read an ODF file.
        /// </summary>
        protected abstract void importOdf();

        /// <summary>
        /// Event handler for Office 2007
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public virtual void ExportOdf12(IRibbonControl control)
        {
            try
            {
                exportOdf();
            }
            catch
            {
            }
        }

        /// <summary>
        /// Event handler for Office versions before 2007
        /// </summary>
        protected virtual void exportButton_Click(object /*CommandBarButton*/ Ctrl, object CancelDefault)
        {
            exportOdf();
        }

        /// <summary>
        /// Save as ODF.
        /// </summary>
        protected abstract void exportOdf();

        /// <summary>
        /// Event handler for Office 2007
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        public virtual void OdfOptions12(IRibbonControl control)
        {
            try
            {
                odfOptions();
            }
            catch
            {
            }
        }

        /// <summary>
        /// Event handler for Office versions before 2007
        /// </summary>
        protected virtual void odfOptionsButton_Click(object /*CommandBarButton*/ Ctrl, object CancelDefault)
        {
            odfOptions();
        }

        protected virtual void odfOptions()
        {
            using (ConfigForm cfgForm = 
                new ConfigForm(this, new System.Resources.ResourceManager("OdfAddinLib.resources.Labels", Assembly.GetExecutingAssembly())))
            {
                cfgForm.ShowDialog();
            }
        }

        #region Office 2007 Members
        /// <summary>
        /// Get custom UI
        /// </summary>
        /// <param name="RibbonID">the ribbon identifier</param>
        /// <returns>string ui</returns>
        string Office.IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            using (System.IO.TextReader tr = new System.IO.StreamReader(getCustomUI()))
            {
                return tr.ReadToEnd();
            }
        }

        /// <summary>
        /// Get an image
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>IPictureDisp object</returns>
        public virtual stdole.IPictureDisp GetImage(Office.IRibbonControl control)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("OdfLogo.png"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }
            }
            if (stream == null)
            {
                return null;
            }
            System.Drawing.Bitmap image = new System.Drawing.Bitmap(stream);
            return ConvertImage.Convert(image);
        }

        /// <summary>
        /// Get a label
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>label as a string</returns>
        public virtual string getLabel(Office.IRibbonControl control)
        {
            return this._addinLib.GetString(control.Id + "Label");
        }

        /// <summary>
        /// Get description
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        /// <returns>Description as a string</returns>
        public virtual string getDescription(Office.IRibbonControl control)
        {
            return this._addinLib.GetString(control.Id + "Description");
        }

        protected virtual Stream getCustomUI()
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
        #endregion

        public abstract string RegistryKeyUser
        {
            get;
        }

        public abstract string RegistryKeyLocalMachine
        {
            get;
        }

        public OdfAddinLib AddinLib
        {
            get
            {
                return _addinLib;
            }
        }

        protected abstract string ImportOdfFileFilter
        {
            get;
        }

        protected virtual string ImportAllFileFilter
        {
            get
            {
                return IMPORT_ALL_FILE_FILTER;
            }
        }

        protected abstract string ExportOdfFileFilter
        {
            get;
        }

        protected abstract string ConnectClassName
        {
            get;
        }

        protected virtual string ExportAllFileFilter
        {
            get
            {
                return EXPORT_ALL_FILE_FILTER;
            }
        }

        protected abstract void InitializeAddin();

        sealed private class ConvertImage : System.Windows.Forms.AxHost
        {
            private ConvertImage()
                : base(null)
            {
            }
            public static stdole.IPictureDisp Convert
                (System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)System.
                    Windows.Forms.AxHost
                    .GetIPictureDispFromPicture(image);
            }
        }
    }
}
