namespace OdfPowerpoint2007Addin
{
    using System;
    using System.Reflection;
    using Microsoft.Office.Core;
    using System.IO;
    using System.Xml;
    using Extensibility;
    using System.Runtime.InteropServices;
    using System.Collections;
    using PPT = Microsoft.Office.Interop.PowerPoint;
    using CleverAge.OdfConverter.OdfConverterLib;
    using Sonata.OdfConverter.OdfPowerpointAddinLib;


	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the OdfPowerpoint2007AddinSetup project, 
	// right click the project in the Solution Explorer, then choose install.
	#endregion
	
	/// <summary>
	///   The object for implementing an Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("D08F6B00-ED78-43E1-8CB1-B83E0386A648"), ProgId("OdfPowerpoint2007Addin.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility, IOdfConverter
	{
        private const string ODF_FILE_TYPE = "OdfFileType";
        private const string ALL_FILE_TYPE = "AllFileType";
        private const string IMPORT_ODF_FILE_FILTER = "*.odp";
        private const string IMPORT_ALL_FILE_FILTER = "*.*";
        private const string IMPORT_LABEL = "OdfImportLabel";
        private const string EXPORT_ODF_FILE_FILTER = " (*.odp)|*.odp|";
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
            this.addinLib = new Sonata.OdfConverter.Presentation.Addin();
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
            this.applicationObject = (PPT.Application)application;
            // set culture to match current application culture or user's choice
            int culture = 0;
            string languageVal = Microsoft.Win32.Registry
                .GetValue(@"HKEY_CURRENT_USER\Software\Sonata\Odf Add-in for Powerpoint", "Language", null) as string;
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
            // add filter for ODP files
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
                object fileName = this.addinLib.GetTempFileName(odfFile, ".pptx");

                // this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;
                OdfToOox(odfFile, (string) fileName, true);
                // this.applicationObject.System.Cursor =  WdCursorType.wdCursorNormal;

                try
                {
                    // conversion may have been cancelled and file deleted.
                    if (File.Exists((string)fileName))
                    {
                        PPT.Presentation p = this.applicationObject.Presentations.Open((string)fileName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                        // p.Activate();
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
        public void ExportODF(IRibbonControl control)
        {
            PPT.Presentation pres = this.applicationObject.ActivePresentation;   
            
            // the second test deals with blank workbooks
            // (which are in a 'saved' state and have no extension yet(?))
            if (pres.Saved == Microsoft.Office.Core.MsoTriState.msoFalse || pres.FullName.IndexOf('.') < 0)
            {
                System.Windows.Forms.MessageBox.Show(addinLib.GetString("OdfSaveDocumentBeforeExport"), DialogBoxTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            }
            else
            {
                System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog();
                // sfd.SupportMultiDottedExtensions = true;
                sfd.AddExtension = true;
                sfd.DefaultExt = "odp";
                sfd.Filter = this.addinLib.GetString(ODF_FILE_TYPE) + EXPORT_ODF_FILE_FILTER
                             + this.addinLib.GetString(ALL_FILE_TYPE) + EXPORT_ALL_FILE_FILTER;
                sfd.InitialDirectory = pres.Path;
                sfd.OverwritePrompt = true;
                sfd.Title = this.addinLib.GetString(EXPORT_LABEL);
                string ext = '.' + sfd.DefaultExt;
                sfd.FileName = pres.FullName.Substring(0, pres.FullName.LastIndexOf('.')) + ext;

                // process the chosen presentations	
                if (System.Windows.Forms.DialogResult.OK == sfd.ShowDialog())
                {
                    // retrieve file name
                    string odfFile = sfd.FileName;
                    if (!odfFile.EndsWith(ext))
                    { // multi dotted extension support
                        odfFile += ext;
                    }
                    object initialName = pres.FullName;
                    object tmpFileName = null;
                    string ooxFile = (string)initialName;

                    //this.applicationObject.System.Cursor = MSExcel.WdCursorType.wdCursorWait;
                 
                    if (!Path.GetExtension(pres.Name).Equals(".pptx"))
                    {
                            // duplicate the file
                            object newName = Path.GetTempFileName() + Path.GetExtension((string)initialName);
                            File.Copy((string)initialName, (string)newName);

                            PPT.Presentation newPres = this.applicationObject.Presentations.Open((string)newName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                            // generate openxml file from the duplicated file (under a temporary file)
                            tmpFileName = Path.GetTempFileName();

                            this.applicationObject.DisplayAlerts = PPT.PpAlertLevel.ppAlertsNone;
                            newPres.SaveAs((string)tmpFileName, PPT.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                            this.applicationObject.DisplayAlerts = PPT.PpAlertLevel.ppAlertsAll;

                            // close and remove the duplicated file
                            newPres.Close();
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
                            ooxFile = (string)tmpFileName;
                        }

                    OoxToOdf(ooxFile, odfFile, true);

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


        private PPT.Application applicationObject;
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