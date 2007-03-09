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
namespace CleverAge.OdfConverter.OdfExcelXPAddin
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
	// you will need to re-register the Add-in by building the OdfExcelXPAddinSetup project, 
	// right click the project in the Solution Explorer, then choose install.
	#endregion
	
	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
    [GuidAttribute("267FE118-B41F-491A-BFE8-9781766BF6F4"), ProgId("OdfExcelXPAddin.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2
    {
        private const int xlOpenXMLWorkbook = 51; // ".xslx" file format
        private string DialogBoxTitle;

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

            this.DialogBoxTitle = addinLib.GetString("OdfConverterTitle");
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
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
            {
                OnBeginShutdown(ref custom);
            }
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
            // Add menu item
            // first retrieve "File" menu
            CommandBar commandBar = applicationObject.CommandBars["File"];

            // Add import button
            try
            {
                // if item already exists, use it (should never happen)
                importButton = (CommandBarButton)commandBar.Controls[this.addinLib.GetString("OdfImportLabel")];
            }
            catch (Exception)
            {
                // otherwise, create a new one
                importButton = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 3, true);
            }
            // set item's label
            importButton.Caption = this.addinLib.GetString("OdfImportLabel");
            importButton.Tag = this.addinLib.GetString("OdfImportLabel");
            // set action
            importButton.OnAction = "!<OdfExcelXPAddin.Connect>";
            importButton.Visible = true;
            importButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.importButton_Click);

            // Add export button
            try
            {
                // if item already exists, use it (should never happen)
                exportButton = (CommandBarButton)commandBar.Controls[this.addinLib.GetString("OdfExportLabel")];
            }
            catch (Exception)
            {
                // otherwise, create a new one
                exportButton = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 4, true);
            }
            // set item's label
            exportButton.Caption = this.addinLib.GetString("OdfExportLabel");
            exportButton.Tag = this.addinLib.GetString("OdfExportLabel");
            // set action
            exportButton.OnAction = "!<OdfExcelXPAddin.Connect>";
            exportButton.Visible = true;
            exportButton.Enabled = true;
            exportButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.exportButton_Click);
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
            // remove menu items (use FindControl instead of referenced object
            // in order to actually remove the items - this is a workaround)
            CommandBarButton button = (CommandBarButton)applicationObject.CommandBars.FindControl(Type.Missing, Type.Missing, this.addinLib.GetString("OdfImportLabel"), Type.Missing);
            button.Delete(Type.Missing);
            button = (CommandBarButton)applicationObject.CommandBars.FindControl(Type.Missing, Type.Missing, this.addinLib.GetString("OdfExportLabel"), Type.Missing);
            button.Delete(Type.Missing);
        }

        /// <summary>
        /// Read an ODF file
        /// </summary>
        /// <param name="control">An IRibbonControl instance</param>
        private void importButton_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
        {
            FileDialog fd = this.applicationObject.get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
            // allow multiple file opening
            fd.AllowMultiSelect = true;
            // add filter for ODS files
            fd.Filters.Clear();
            fd.Filters.Add(this.addinLib.GetString("OdfFileType"), "*.ods", Type.Missing);
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
                object fileName = this.addinLib.GetTempFileName(odfFile, ".xlsx");

               // this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;
                OdfToOox(odfFile, (string)fileName, true);
                // this.applicationObject.System.Cursor =  WdCursorType.wdCursorNormal;

                try
                {
                    // open the workbook
                    object readOnly = true;
                    object addToRecentFiles = false;
                    object isVisible = true;
                    object openAndRepair = false;
                    object missing = Missing.Value;

                    // conversion may have been cancelled and file deleted.
                    if (File.Exists((string)fileName))
                    {
                        // Workaround to excel bug. "Old format or invalid type library. (Exception from HRESULT: 0x80028018 (TYPE_E_INVDATAREAD))" 
                        System.Globalization.CultureInfo ci;
                        ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                   
                        Microsoft.Office.Interop.Excel.Workbook wb =
                            this.applicationObject.Workbooks.Open((string) fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                        
                        wb.Activate();
                       
                        System.Threading.Thread.CurrentThread.CurrentCulture = ci;
                    }
                }
                catch (Exception ex)
                {
                    // this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorNormal;
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
        private void exportButton_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
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
                    sfd.Filter = this.addinLib.GetString("OdfFileType") + " (*.ods)|*.ods|"
                            + this.addinLib.GetString("AllFileType") + " (*.*)|*.*";
                    sfd.InitialDirectory = wb.Path;
                    sfd.OverwritePrompt = true;
                    sfd.Title = this.addinLib.GetString("OdfExportLabel");
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
                        object tmpFileName = null;
                        string xlsxFile = (string)initialName;

                        //this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;

                        if (!"50".Equals(wb.FileFormat.ToString()))
                        {
                            // duplicate the file
                            object newName = Path.GetTempFileName() + Path.GetExtension((string)initialName);
                            File.Copy((string)initialName, (string)newName);

                            // open the duplicated file
                            object addToRecentFiles = false;
                            object readOnly = false;
                            object isVisible = false;
                            object missing = Missing.Value;

                            MSExcel.Workbook newWb = this.applicationObject.Workbooks.Open((string)newName, missing, readOnly, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                            // generate xlsx file from the duplicated file (under a temporary file)
                            tmpFileName = Path.GetTempFileName();
                            try
                            {
                                this.applicationObject.DisplayAlerts = false;
                                newWb.SaveAs((string)tmpFileName, xlOpenXMLWorkbook, missing, missing, missing, missing, MSExcel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                            }
                            catch (Exception e)
                            {
                                // xlOpenXMLWorkbook file format not supported. Compatibility pack not installed.
                                System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfConverterNotInstalled"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                                return;
                            }
                            finally
                            {
                                this.applicationObject.DisplayAlerts = true;
                                // close and remove the duplicated file
                                newWb.Close(false, false, missing);
                                try
                                {
                                    File.Delete((string)newName);
                                }
                                catch (IOException)
                                {
                                    // bug #1610099
                                    // deletion failed : file currently used by another application.
                                    // do nothing
                                }
                            }
                            xlsxFile = (string)tmpFileName;
                        }

                        OoxToOdf(xlsxFile, odfFile, true);

                        if (tmpFileName != null && File.Exists((string)tmpFileName))
                        {
                            File.Delete((string)tmpFileName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Need to catch a file format exception here.
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            }
        }


        private MSExcel.Application applicationObject;
        private OdfAddinLib addinLib;
        private CommandBarButton importButton, exportButton;

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