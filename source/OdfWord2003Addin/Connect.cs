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

namespace CleverAge.OdfConverter.OdfWord2003Addin
{
    using System;
    using System.Reflection;
    using Microsoft.Office.Core;
    using System.IO;
    using System.Xml;
    using Extensibility;
    using System.Runtime.InteropServices;
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
    [GuidAttribute("CA43CC65-0A7F-4CA0-97A8-3BA6BFA7AAFB"), ProgId("OdfWord2003Addin.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2
    {
        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>
        public Connect()
        {
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
            this.labelsResourceManager = OdfWordAddinLib.GetResourceManager();

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
            this.labelsResourceManager = null;
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
                importButton = (CommandBarButton)commandBar.Controls[labelsResourceManager.GetString("OdfImportLabel")];
            }
            catch (Exception)
            {
                // otherwise, create a new one
                importButton = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 3, true);
            }
            // set item's label
            importButton.Caption = labelsResourceManager.GetString("OdfImportLabel");
            importButton.Tag = labelsResourceManager.GetString("OdfImportLabel");
            // set action
            importButton.OnAction = "!<OpenOfficePlugin.Connect>";
            importButton.Visible = true;
            importButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.importButton_Click);

            // Add export button
            try
            {
                // if item already exists, use it (should never happen)
                exportButton = (CommandBarButton)commandBar.Controls[labelsResourceManager.GetString("OdfExportLabel")];
            }
            catch (Exception)
            {
                // otherwise, create a new one
                exportButton = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 4, true);
            }
            // set item's label
            exportButton.Caption = labelsResourceManager.GetString("OdfExportLabel");
            exportButton.Tag = labelsResourceManager.GetString("OdfExportLabel");
            // set action
            exportButton.OnAction = "!<OpenOfficePlugin.Connect>";
            exportButton.Visible = true;
            exportButton.Enabled = false;
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
            CommandBarButton button = (CommandBarButton)applicationObject.CommandBars.FindControl(Type.Missing, Type.Missing, labelsResourceManager.GetString("OdfImportLabel"), Type.Missing);
            button.Delete(Type.Missing);
            button = (CommandBarButton)applicationObject.CommandBars.FindControl(Type.Missing, Type.Missing, labelsResourceManager.GetString("OdfExportLabel"), Type.Missing);
            button.Delete(Type.Missing);
            applicationObject.NormalTemplate.Save();
        }

        private void importButton_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
        {
            FileDialog fd = applicationObject.get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
            // allow multiple file opening
            fd.AllowMultiSelect = true;
            // add filter for ODT files
            fd.Filters.Clear();
            fd.Filters.Add(labelsResourceManager.GetString("OdfFileType"), "*.odt", Type.Missing);
            fd.Filters.Add(labelsResourceManager.GetString("AllFileType"), "*.*", Type.Missing);
            // set title
            fd.Title = labelsResourceManager.GetString("OdfImportLabel");
            // display the dialog
            fd.Show();

            // process the chosen documents	
            for (int i = 1; i <= fd.SelectedItems.Count; ++i)
            {
                // retrieve file name
                String odfFile = fd.SelectedItems.Item(i);

                // instanciate tranformer and process the file
                object fileName = null;
                ConverterForm form = null;

                try
                {
                    applicationObject.System.Cursor = MSword.WdCursorType.wdCursorWait;
                    // create a temporary file
                    fileName = OdfWordAddinLib.GetTempFileName(odfFile);

                    // call the converter
                    using (form = new ConverterForm(odfFile, (string)fileName, labelsResourceManager))
                    {
                        System.Windows.Forms.DialogResult dr = form.ShowDialog();

                        if (form.Exception != null)
                        {
                            throw form.Exception;
                        }

                        if (dr == System.Windows.Forms.DialogResult.OK)
                        {
                            // open the document
                            object readOnly = true;
                            object isVisible = true;
                            object Format = MSword.WdOpenFormat.wdOpenFormatXML;
                            object missing = Type.Missing;
                            Microsoft.Office.Interop.Word.Document doc = applicationObject.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

                            // and activate it
                            doc.Activate();
                        }
                        else
                        {
                            if (File.Exists((string)fileName))
                            {
                                File.Delete((string)fileName);
                            }
                        }
                    }
                }
                catch (NotAnOdfDocumentException)
                {
                    System.Windows.Forms.MessageBox.Show(odfFile + " " + labelsResourceManager.GetString("NotAnOdfDocumentError"));
                }
                catch (Exception e)
                {
                    if (File.Exists((string)fileName))
                    {
                        File.Delete((string)fileName);
                    }
#if DEBUG
                    System.Windows.Forms.MessageBox.Show(e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");
#else
                    System.Windows.Forms.MessageBox.Show(labelsResourceManager.GetString("OdfUnexpectedError"));
#endif
                }
                finally
                {
                    applicationObject.System.Cursor = MSword.WdCursorType.wdCursorNormal;
                }

            }
        }

        private void exportButton_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
        {
            System.Windows.Forms.MessageBox.Show(labelsResourceManager.GetString("NotImplemented"));
        }

        private MSword.Application applicationObject;
        private System.Resources.ResourceManager labelsResourceManager;
        private CommandBarButton importButton, exportButton;
    }
}
