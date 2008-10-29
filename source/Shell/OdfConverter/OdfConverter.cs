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

using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Threading;

using CleverAge.OdfConverter.OdfZipUtils;
using CleverAge.OdfConverter.OdfConverterLib;
using Sonata.OdfConverter.Presentation;
using CleverAge.OdfConverter.Spreadsheet;
using Wordprocessing = OdfConverter.Wordprocessing;
using System.Collections.Generic;


namespace CleverAge.OdfConverter.CommandLineTool
{
    //enum ControlType : int
    //{
    //    CTRL_C_EVENT = 0,
    //    CTRL_BREAK_EVENT = 1,
    //    CTRL_CLOSE_EVENT = 2,
    //    CTRL_LOGOFF_EVENT = 5,
    //    CTRL_SHUTDOWN_EVENT = 6
    //};

    //delegate int ControlHandlerFunction(ControlType control);

    class Word
    {
        Type _type;
        object _instance;
        Type _docsType;
        object _documents;

        public Word()
        {
            _type = Type.GetTypeFromProgID("Word.Application");
            _instance = Activator.CreateInstance(_type);
            _docsType = null;
            _documents = null;
        }

        public bool Visible
        {
            set
            {
                object[] args = new object[] { value };
                _type.InvokeMember("Visible", BindingFlags.SetProperty, null, _instance, args);
            }
        }

        public void Quit()
        {
            object[] args = new object[] { Missing.Value,
                                            Missing.Value,
                                            Missing.Value };
            _type.InvokeMember("Quit", BindingFlags.InvokeMethod, null, _instance, args);
        }

        public void Open(string document)
        {
            if (_documents == null)
            {
                _documents = _type.InvokeMember("Documents", BindingFlags.GetProperty, null, _instance, null);
                _docsType = _documents.GetType();
            }
            object[] args = new object[] { document };
            _docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, _documents, args);

        }
    }

    class Excel
    {
        Type _type;
        object _instance;
        Type _docsType;
        object _documents;

        public Excel()
        {
            _type = Type.GetTypeFromProgID("Excel.Application");
            _instance = Activator.CreateInstance(_type);
            _docsType = null;
            _documents = null;
        }

        public bool Visible
        {
            set
            {
                object[] args = new object[] { value };
                _type.InvokeMember("Visible", BindingFlags.SetProperty, null, _instance, args);
            }
        }

        public void Quit()
        {
            object[] args = new object[] { Missing.Value,
                                            Missing.Value,
                                            Missing.Value };
            _type.InvokeMember("Quit", BindingFlags.InvokeMethod, null, _instance, args);
        }

        public void Open(string document)
        {
            if (_documents == null)
            {
                _documents = _type.InvokeMember("Documents", BindingFlags.GetProperty, null, _instance, null);
                _docsType = _documents.GetType();
            }
            object[] args = new object[] { document };
            _docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, _documents, args);

        }
    }

    class Presentation
    {
        Type _type;
        object _instance;
        Type _docsType;
        object _documents;

        public Presentation()
        {
            _type = Type.GetTypeFromProgID("POWERPNT.Application");
            _instance = Activator.CreateInstance(_type);
            _docsType = null;
            _documents = null;
        }

        public bool Visible
        {
            set
            {
                object[] args = new object[] { value };
                _type.InvokeMember("Visible", BindingFlags.SetProperty, null, _instance, args);
            }
        }

        public void Quit()
        {
            object[] args = new object[] { Missing.Value,
                                            Missing.Value,
                                            Missing.Value };
            _type.InvokeMember("Quit", BindingFlags.InvokeMethod, null, _instance, args);
        }

        public void Open(string document)
        {
            if (_documents == null)
            {
                _documents = _type.InvokeMember("Documents", BindingFlags.GetProperty, null, _instance, null);
                _docsType = _documents.GetType();
            }
            object[] args = new object[] { document };
            _docsType.InvokeMember("Open", BindingFlags.InvokeMethod, null, _documents, args);

        }
    }

    /// <summary>
    /// OdfConverter is a CommandLine wrapper for the OpenXML - ODF translators
    /// 
    /// Execute the command without argument to see the options.
    /// </summary>
    public class OdfConverter
    {
        private ConversionOptions _options = new ConversionOptions();
        
        //private string input = null;                     // input path
        //private string output = null;                    // output path
        //private bool validate = false;                   // validate the result of the transformations
        //private bool open = false;                       // try to open the result of the transformations
        //private bool recursiveMode = false;              // go in subfolders ?
        //private bool replace = false;					   // override existing files ?
        //private string reportPath = null;                // file to save report
        //private int reportLevel = ConversionReport.INFO_LEVEL;     // file to save report
        //private string xslPath = null;                   // Path to an external stylesheet
        //private ArrayList skippedPostProcessors = null;  // Post processors to skip (identified by their names)
        //private bool packaging = true;                   // Build the zip archive after conversion

	   	//private Direction transformDirection = Direction.OdtToDocx; // direction of conversion
		private bool transformDirectionOverride = false; // whether conversion direction has been specified
        private ConversionReport report = null;
        private Word word = null;
        private Excel excel = null;
        private Presentation presentation = null; 
        private OoxValidator ooxValidator = null;
        private OdfValidator odfValidator = null;
        private Direction batch = Direction.None;

//#if MONO
//        static bool SetConsoleCtrlHandler(ControlHandlerFunction handlerRoutine, bool add) 
//        { return true; }
//#else
//        [DllImport("kernel32")]
//        static extern bool SetConsoleCtrlHandler(ControlHandlerFunction handlerRoutine, bool add);
//#endif

//        int MyHandler(ControlType control)
//        {
//            //Console.WriteLine("MyHandler: " + control.ToString());
//            if (word != null)
//            {
//                try
//                {
//                    word.Quit();
//                }
//                catch
//                {
//                    Console.WriteLine("Unable to close Word, please close it manually");
//                }
//            }
//            return 0;
//        }

        /// <summary>
        /// Main program.
        /// </summary>
        /// <param name="args">Command Line arguments</param>
        public static void Main(String[] args)
        {
            // math, dialogika: Comment this out as we need the culture info for the list of TOC.
            // Moreover, changing the culture info only for the command-line-tool changes the
            // behaviour between add-in and command-line-tool which makes detecting and fixing bugs more difficult
            // System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
            OdfConverter tester = new OdfConverter();
            //ControlHandlerFunction myHandler = new ControlHandlerFunction(tester.MyHandler);
            //SetConsoleCtrlHandler(myHandler, true);
            try
            {
                tester.ParseCommandLine(args);
            }
            catch (Exception e)
            {
            	Environment.ExitCode = 1;
                Console.WriteLine("Error when parsing command line: " + e.Message);
                usage();
                return;
            }
            try
            {
                tester.proceed();
                Console.WriteLine("Done.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception raised when running conversion: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private OdfConverter()
        {
        }

        protected ConversionOptions Options
        {
            get { return _options; }
            set { _options = value; }
        }

        private void proceed()
        {
            this.report = new ConversionReport(this.Options.ReportPath, this.Options.ReportLevel);

            switch (this.batch)
            {
                case Direction.OdsToXlsx:
                case Direction.OdpToPptx:
                case Direction.OdtToDocx:
                    this.ProceedBatchOdf();
                    break;
                case Direction.DocxToOdt:
                case Direction.PptxToOdp:
                case Direction.XlsxToOds:
                    this.ProceedBatchOox();
                    break;
                default:  // no batch mode
                    // instanciate Word if needed
                    if (this.Options.TransformDirection == Direction.OdtToDocx && this.Options.Open)
                    {
                        word = new Word();
                        word.Visible = false;
                    }

                    this.ProceedSingleFile(this.Options.InputPath, this.Options.OutputPath, this.Options.TransformDirection);

                    // close word if needed
                    if (this.Options.Open)
                    {
                        word.Quit();
                    }
                    break;
            }

            this.report.Close();
        }

        private void ProceedBatchOdf()
        {
            // instanciate word if needed
            if (this.Options.Open && (this.batch == Direction.OdtToDocx))
            {
                this.word = new Word();
                this.word.Visible = false;
            }
            
            // instanciate excel if needed
            if (this.Options.Open && (this.batch == Direction.OdsToXlsx))
            {
                
                this.excel = new Excel();
                this.excel.Visible = false;
            }
           
            // instanciate Presentation if needed
            if (this.Options.Open && (this.batch == Direction.OdpToPptx))
            {
                this.presentation = new Presentation();
                this.presentation.Visible = false;
            }

            // instanciate validator if needed
            if (this.Options.Validate)
            {
                this.report.AddComment("Instanciating validator, please wait...");
                this.ooxValidator = new OoxValidator(this.report);
                this.report.AddComment("Validator instanciated");
            }

            string pattern = "";
            string targetExt = null;

            switch (this.batch)
            {
                case Direction.OdtToDocx:
                    pattern = "*.odt;*.ott";
                    targetExt = ".docx";
                    break;
                case Direction.OdsToXlsx:
                    pattern = "*.ods";
                    targetExt = ".xlsx";
                    break;
                case Direction.OdpToPptx:
                    pattern = "*.odp";
                    targetExt = ".pptx";
                    break;
                default:
                    throw new ArgumentException("unsupported batch type");
            }

            string[] files = GetFiles(this.Options.InputPath, pattern, this.Options.RecursiveMode);

            int nbFiles = files.Length;
            int nbConverted = 0;
            int nbValidatedAndOpened = 0;
            int nbValidatedAndNotOpened = 0;
            int nbNotValidatedAndOpened = 0;
            int nbNotValidatedAndNotOpened = 0;
            this.report.AddComment("Processing " + nbFiles + " ODF file(s)");
            foreach (string input in files)
            {
                string output = this.GenerateOutputName(this.Options.OutputPath, input, targetExt, this.Options.ForceOverwrite);
                int result = this.ProceedSingleFile(input, output, this.batch);
                switch (result)
                {
                    case NOT_CONVERTED:
                        break;
                    case VALIDATED_AND_OPENED:
                        nbConverted++;
                        nbValidatedAndOpened++;
                        break;
                    case VALIDATED_AND_NOT_OPENED:
                        nbConverted++;
                        nbValidatedAndNotOpened++;
                        break;
                    case NOT_VALIDATED_AND_OPENED:
                        nbConverted++;
                        nbNotValidatedAndOpened++;
                        break;
                    case NOT_VALIDATED_AND_NOT_OPENED:
                        nbConverted++;
                        nbNotValidatedAndNotOpened++;
                        break;
                    default:
                        break;
                }
            }
            // close word if needed
            if (this.Options.Open)
            {
                word.Quit();
            }
            //string varResult = null;
            //if (ext == "odp" || ext=="pptx")
            //{
            //    varResult = "PowerPoint";
            //}
            //if (ext == "ods" || ext == "xlsx")
            //{
            //    varResult = "Excel";
            //}
            //if (ext == "odt" || ext == "docx")
            //{
            //    varResult = "Word";
            //}

            this.report.AddComment("Results: " + nbConverted + " file(s) over " + nbFiles + " were converted successfully.");

            // opening and validating options have been removed
            //this.report.AddComment("Results: " + nbConverted + " file(s) over " + nbFiles + " were converted, among them:");
            //this.report.AddComment("   " + nbValidatedAndOpened + " were validated and successfully opened in " + varResult);
            //this.report.AddComment("   " + nbValidatedAndNotOpened + " were validated but could not be opened in " + varResult);
            //this.report.AddComment("   " + nbNotValidatedAndOpened + " were not validated but were successfully opened  " + varResult);
            //this.report.AddComment("   " + nbNotValidatedAndNotOpened + " were not validated and could not be opened in " + varResult);
        }

        private void ProceedBatchOox()
        {
            string pattern;
            string targetExt;
            switch (this.batch)
            {
                case Direction.DocxToOdt:
                    pattern = "*.docx;*.docm;*.dotx;*.dotm";
                    targetExt = ".odt";
                    break;
                case Direction.XlsxToOds:
                    pattern = "*.xlsx";
                    targetExt = ".ods";
                    break;
                case Direction.PptxToOdp:
                    pattern = "*.pptx";
                    targetExt = ".odp";
                    break;
                default:
                    throw new ArgumentException("wrong batch mode");
            }
            // instantiate validator if needed
            if (this.Options.Validate)
            {
                this.report.AddComment("Instantiating validator, please wait...");
                this.odfValidator = new OdfValidator(this.report);
                this.report.AddComment("Validator instanciated");
            }

            string[] files = GetFiles(this.Options.InputPath, pattern, this.Options.RecursiveMode);
            foreach (string input in files)
            {
                // do not process temp files of Office
                if (!Path.GetFileName(input).StartsWith("~$"))
                {
                    string output = this.GenerateOutputName(this.Options.OutputPath, input, targetExt, this.Options.ForceOverwrite);
                    this.ProceedSingleFile(input, output, this.batch);
                }
            }
        }

        /// <summary>
        /// Retrieve a list of files from a folder fulfilling a certain search pattern
        /// </summary>
        /// <param name="path">The directory to search</param>
        /// <param name="searchPattern">A ;-separated list of wildcard patterns to match against the files in <paramref name="path"/></param>
        /// <param name="recursive">If true, the search includes all subdirectories, otherwise only the only the current directory searched.</param>
        /// <returns>A String array containing the names of files in the specified directory that 
        ///     match the specified search pattern. File names include the full path. </returns>
        private string[] GetFiles(string path, string searchPattern, bool recursive)
        {
            string[] m_arExt = searchPattern.Split(';');

            SearchOption option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            List<string> strFiles = new List<string>();
            foreach (string filter in m_arExt)
            {
                strFiles.AddRange(System.IO.Directory.GetFiles(path, filter, option));
            }
            return strFiles.ToArray();
        }

        private int ProceedSingleFile(string input, string output, Direction transformDirection)
        {
            bool converted = false;
            bool validated = false;
            bool opened = false;
            report.AddLog(input, "Converting file: " + input + " into " + output, ConversionReport.INFO_LEVEL);
            converted = ConvertFile(input, output, transformDirection);
            if (converted && this.Options.Validate)
            {
                validated = ValidateFile(input, output, transformDirection);
            }
            if (converted && this.Options.Open)
            {
                opened = TryToOpen(input, output, transformDirection);
            }
            if (!converted) 
            {
            	Environment.ExitCode = 1;
            	return NOT_CONVERTED;
            }
            else if (validated && opened) return VALIDATED_AND_OPENED;
            else if (validated && !opened) return VALIDATED_AND_NOT_OPENED;
            else if (!validated && opened) return NOT_VALIDATED_AND_OPENED;
            else return NOT_VALIDATED_AND_NOT_OPENED;
        }

        private bool ConvertFile(string input, string output, Direction transformDirection)
        {
           
            try
            {
                DateTime start = DateTime.Now;
                AbstractConverter converter = ConverterFactory.Instance(transformDirection);
                converter.ExternalResources = this.Options.XslPath;
                converter.SkippedPostProcessors = this.Options.SkippedPostProcessors;
                converter.DirectTransform =
                    transformDirection == Direction.OdtToDocx
                    || transformDirection == Direction.OdpToPptx
                    || transformDirection == Direction.OdsToXlsx;
                converter.Packaging = this.Options.Packaging;

                // TODO: optionally we could also add a progress bar to the command line
                // then we would need to count the paragraphs, runs, etc.
                //converter.ComputeSize(input);
                
                converter.Transform(input, output);
                TimeSpan duration = DateTime.Now - start;
                this.report.AddLog(input, "Conversion succeeded", ConversionReport.INFO_LEVEL);
                this.report.AddLog(input, "Total conversion time: " + duration, ConversionReport.INFO_LEVEL);
                return true;
            }
            catch (EncryptedDocumentException)
            {
                this.report.AddLog(input, "Conversion failed - Input file is encrypted", ConversionReport.WARNING_LEVEL);
                return false;
            }
            catch (NotAnOdfDocumentException e)
            {
                this.report.AddLog(input, "Conversion failed - Input file is not a valid ODF file", ConversionReport.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                return false;
            }
            catch (NotAnOoxDocumentException e)
            {
                this.report.AddLog(input, "Conversion failed - Input file is not a valid Office OpenXML file", ConversionReport.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                return false;
            }
            catch (ZipException e)
            {
                this.report.AddLog(input, "Conversion failed - Input file is not a valid file for conversion or might be password protected", ConversionReport.ERROR_LEVEL);
                this.report.AddLog(input, e.Message, ConversionReport.DEBUG_LEVEL);
                return false;
            }
            //Pradeep Nemadi - Bug 1747083 Start
            //IOException is added to fix this bug
            catch (IOException e)
            {
                this.report.AddLog(input, "Conversion failed - " + e.Message, ConversionReport.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                return false;
            }
            //Pradeep Nemadi - Bug 1747083 end
            catch (Exception e)
            {
                this.report.AddLog(input, "Conversion failed - Error during conversion", ConversionReport.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                return false;
            }
        }

        private bool ValidateFile(string input, string output, Direction transformDirection)
        {
            if (transformDirection == Direction.OdtToDocx 
                || transformDirection == Direction.OdsToXlsx
                || transformDirection == Direction.OdpToPptx)
            {
                try
                {
                    if (this.ooxValidator == null)
                    {
                        this.report.AddComment("Instanciating validator, please wait...");
                        this.ooxValidator = new OoxValidator(this.report);
                        this.report.AddComment("Validator instanciated");
                    }
                    this.ooxValidator.validate(output);
                    this.report.AddLog(input, "Converted file (" + output + ") is valid", ConversionReport.INFO_LEVEL);
                    return true;
                }
                catch (OoxValidatorException e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") is invalid", ConversionReport.WARNING_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "An unexpected exception occurred when trying to validate " + output, ConversionReport.ERROR_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
            }
            else
            {
                try
                {
                    if (this.odfValidator == null)
                    {
                        this.report.AddComment("Instanciating validator, please wait...");
                        this.odfValidator = new OdfValidator(this.report);
                        this.report.AddComment("Validator instanciated");
                    }
                    this.odfValidator.validate(output);
                    this.report.AddLog(input, "Converted file (" + output + ") is valid", ConversionReport.INFO_LEVEL);
                    return true;
                }
                catch (OdfValidatorException e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") is invalid", ConversionReport.WARNING_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "An unexpected exception occurred when trying to validate " + output, ConversionReport.ERROR_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
            }
        }

        private bool TryToOpen(string input, string output, Direction transformDirection)
        {
            if (transformDirection == Direction.OdtToDocx)
            {
                //  Microsoft.Office.Interop.Word.Application msWord = (Microsoft.Office.Interop.Word.Application)wordApp;
                try
                {
                    string filename = Path.GetFullPath(output);
                    word.Open(filename);
                    this.report.AddLog(input, "Converted file opened successfully in Word", ConversionReport.INFO_LEVEL);
                    return true;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") could not be opened in Word", ConversionReport.ERROR_LEVEL);
                    this.report.AddLog(input, e.GetType().Name + ": " + e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
            }

            if (transformDirection == Direction.OdpToPptx)
            {
                try
                {
                    string filename = Path.GetFullPath(output);
                    presentation.Open(filename);
                    this.report.AddLog(input, "Converted file opened successfully in PowerPoint", ConversionReport.INFO_LEVEL);
                    return true;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") could not be opened in PowerPoint", ConversionReport.ERROR_LEVEL);
                    this.report.AddLog(input, e.GetType().Name + ": " + e.Message + "(" + e.StackTrace + ")", ConversionReport.DEBUG_LEVEL);
                    return false;
                }
            }
            return false;
        }

        private static void usage()
        {
            Console.WriteLine("Usage: OdfConverter.exe /I PathOrFilename [/O PathOrFilename] [/BATCH-ODT] [/BATCH-DOCX] [/XSLT Path] [/NOPACKAGING] [/SKIP Name] [/REPORT Filename] [/LEVEL Level] ");
            //Console.WriteLine("Usage: OdfConverter.exe /I PathOrFilename [/O PathOrFilename] [/BATCH-ODT] [/BATCH-DOCX] [/V] [/OPEN] [/XSLT Path] [/NOPACKAGING] [/SKIP name] [/REPORT Filename] [/LEVEL Level] ");
            Console.WriteLine("  Where options are:");
            Console.WriteLine("     /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)");
            Console.WriteLine("     /O PathOrFilename  Name of the output file (or output folder)");
            Console.WriteLine("     /F                 Replace existing file");
            Console.WriteLine("     /BATCH-ODT         Do a batch conversion over every ODT file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-ODP         Do a batch conversion over every ODP file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-PPTX        Do a batch conversion over every PPTX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-ODS         Do a batch conversion over every ODS file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-XLSX        Do a batch conversion over every XLSX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /R                 Process subfolders recursively during batch conversion");
            //Console.WriteLine("     /V                 Validate the result of the transformation against the schemas");
            //Console.WriteLine("     /OPEN              Try to open the converted files (works only for ODF->OOX, Microsoft Word required)");
            Console.WriteLine("     /XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)");
            Console.WriteLine("     /NOPACKAGING       Don't package the result of the transformation into a ZIP archive (produce raw XML)");
            Console.WriteLine("     /SKIP Name         Skip a post-processing (provide the post-processor's name)");
            Console.WriteLine("     /REPORT Filename   Name of the report file that must be generated (existing files will be replaced)");
            Console.WriteLine("     /LEVEL Level       Level of reporting: 1=DEBUG, 2=INFO, 3=WARNING, 4=ERROR");
          	Console.WriteLine("     /ODT2DOCX          Force conversion to DOCX regardless of input file extension");
			Console.WriteLine("     /DOCX2ODT          Force conversion to ODT regardless of input file extension");
            Console.WriteLine("     /ODS2XLSX          Force conversion to XLSX regardless of input file extension");
            Console.WriteLine("     /XLSX2ODS          Force conversion to ODS regardless of input file extension");
            Console.WriteLine("     /ODP2PPTX          Force conversion to PPTX regardless of input file extension");
            Console.WriteLine("     /PPTX2ODP          Force conversion to ODP regardless of input file extension");
        }

        private void ParseCommandLine(string[] args)
        {
            this.RetrieveArgs(args);
            this.CheckPaths();
        }

        private void RetrieveArgs(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                // FIX: 20071030/divo: accept lower case parameters
                switch (args[i].ToUpperInvariant())
                {
                    case "/I":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Input missing");
                        }
                        this.Options.InputPath = args[i];
                        break;
                    case "/O":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Output missing");
                        }
                        this.Options.OutputPath = args[i];
                        break;
                    //case "/V":
                    //    this.validate = true;
                    //    break;
                    //case "/OPEN":
                    //    this.open = true;
                    //    break;
                    case "/LEVEL":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Level missing");
                        }
                        try
                        {
                            this.Options.ReportLevel = int.Parse(args[i]);
                        }
                        catch (Exception)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1,2 3 or 4)");
                        }
                        if (this.Options.ReportLevel < 1 || this.Options.ReportLevel > 4)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1,2 3 or 4)");
                        }
                        break;
                    case "/R":
                        this.Options.RecursiveMode = true;
                        break;
                    case "/NOPACKAGING":
                        this.Options.Packaging = false;
                        break;
                    case "/BATCH-ODT":
                        this.batch = Direction.OdtToDocx; 
                        break;
                    case "/BATCH-DOCX":
                        this.batch = Direction.DocxToOdt;
                        break;
                    case "/BATCH-ODP":
                        this.batch = Direction.OdpToPptx;
                        break;
                    case "/BATCH-PPTX":
                        this.batch = Direction.PptxToOdp;
                        break;
                    case "/BATCH-ODS":
                        this.batch = Direction.OdsToXlsx;
                        break;
                    case "/BATCH-XLSX":
                        this.batch = Direction.XlsxToOds;
                        break;
                    case "/SKIP":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Post processing name missing");
                        }
                        this.Options.SkippedPostProcessors.Add(args[i]);
                        break;
                    case "/REPORT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Report file missing");
                        }
                        this.Options.ReportPath = args[i];
                        break;
                    case "/XSLT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("XSLT path missing");
                        }
                        this.Options.XslPath = args[i];
                        break;
                    case "/CV":
                        OoxValidator.test();
                        break;
                   	case "/DOCX2ODT":
                        this.Options.TransformDirection = Direction.DocxToOdt;
						this.transformDirectionOverride = true;
						// System.Console.WriteLine("Override to odt\n");
						break;
					case "/ODT2DOCX":
                        this.Options.TransformDirection = Direction.OdtToDocx;
						this.transformDirectionOverride = true;
						// System.Console.WriteLine("Override to docx\n");
						break;
                    case "/XLSX2ODS":
                        this.Options.TransformDirection = Direction.XlsxToOds;
                        this.transformDirectionOverride = true;
                        // System.Console.WriteLine("Override to odt\n");
                        break;
                    case "/ODS2XLSX":
                        this.Options.TransformDirection = Direction.OdsToXlsx;
                        this.transformDirectionOverride = true;
                        // System.Console.WriteLine("Override to docx\n");
                        break;
                    case "/PPTX2ODP":
                        this.Options.TransformDirection = Direction.PptxToOdp;
                        this.transformDirectionOverride = true;
                        break;
                    case "/ODP2PPTX":
                        this.Options.TransformDirection = Direction.OdpToPptx;
                        this.transformDirectionOverride = true;
                        break;
					case "/F":
						this.Options.ForceOverwrite = true;
						break;
                    default:
                        if (string.IsNullOrEmpty(this.Options.InputPath))
                        {
                            this.Options.InputPath = args[i];
                        }
                        break;
                }
            }
        }

        private void CheckPaths()
        {
            if (string.IsNullOrEmpty(this.Options.InputPath))
            {
                throw new OdfCommandLineException("Input is missing");
            }
            if (this.batch != Direction.None)
            {
                this.CheckBatch();
            }
            else
            {
                this.CheckSingleFile();
            }
        }

        private void CheckBatch()
        {
            if (!Directory.Exists(this.Options.InputPath))
            {
                throw new OdfCommandLineException("Input folder does not exist");
            }
            if (File.Exists(this.Options.OutputPath))
            {
                throw new OdfCommandLineException("Output must be a folder");
            }
            if (string.IsNullOrEmpty(this.Options.OutputPath))
            {
                // use input folder as output folder
                this.Options.OutputPath = this.Options.InputPath;
            }
            if (!Directory.Exists(this.Options.OutputPath))
            {
                try
                {
                    Directory.CreateDirectory(this.Options.OutputPath);
                }
                catch (Exception)
                {
                    throw new OdfCommandLineException(string.Format("Cannot create output folder {0}", this.Options.OutputPath));
                }
            }
        }

        private void CheckSingleFile()
        {
            if (!File.Exists(this.Options.InputPath))
            {
                throw new OdfCommandLineException("Input file does not exist");
            }

            string extension = null;
            if (transformDirectionOverride)
            {
                switch (this.Options.TransformDirection) 
                {
                    case Direction.DocxToOdt:
                        extension = ".odt";
                        break;
                    case Direction.OdtToDocx:
                        extension = ".docx";
                        break;
                    case Direction.PptxToOdp:
                        extension = ".odp";
                        break;
                    case Direction.OdpToPptx:
                        extension = ".pptx";
                        break;
                    case Direction.OdsToXlsx:
                        extension = ".xlsx";
                        break;
                    case Direction.XlsxToOds:
                        extension = ".ods";
                        break;
                }
            }
            else
            {
                // auto-detect transform direction based on file extension
                extension = ".docx";
                string inputExtension = Path.GetExtension(this.Options.InputPath).ToLowerInvariant();
                string outputExtension = "";
                if (!string.IsNullOrEmpty(this.Options.OutputPath))
                {
                    FileInfo fi = new FileInfo(this.Options.OutputPath);
                    outputExtension = fi.Extension.ToLowerInvariant();
                }
                switch (inputExtension)
                {
                    case ".odt":
                    case ".ott":
                        if (outputExtension.Equals(".docx") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.OdtToDocx;
                            extension = ".docx";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        if (outputExtension.Equals(".odt") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.DocxToOdt;
                            extension = ".odt";
                        }
                        else if (outputExtension.Equals(".ott"))
                        {
                            this.Options.TransformDirection = Direction.DocxToOdt;
                            extension = ".ott";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    case ".odp":
                        if (outputExtension.Equals(".pptx") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.OdpToPptx;
                            extension = ".pptx";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    case ".pptx":
                        if (outputExtension.Equals(".odp") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.PptxToOdp;
                            extension = ".odp";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    case ".ods":
                        if (outputExtension.Equals(".xlsx") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.OdsToXlsx;
                            extension = ".xlsx";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    case ".xlsx":
                        if (outputExtension.Equals(".ods") || outputExtension.Equals(""))
                        {
                            this.Options.TransformDirection = Direction.XlsxToOds;
                            extension = ".ods";
                        }
                        else
                        {
                            throw new OdfCommandLineException("Output file extension is invalid");
                        }
                        break;
                    default:
                        throw new OdfCommandLineException("Input file extension [" + inputExtension + "] is not supported.");
                }
            }
            if (!this.Options.Packaging)
            {
                extension = ".xml";
            }

            if (!File.Exists(this.Options.OutputPath) && (this.Options.OutputPath == null))
            {
                string outputPath = this.Options.OutputPath;
                if (outputPath == null)
                {
                    // we take input path
                    outputPath = Path.GetDirectoryName(this.Options.InputPath);
                }
                this.Options.OutputPath = GenerateOutputName(outputPath, this.Options.InputPath, extension, this.Options.ForceOverwrite);
            }
        }

        private string GenerateOutputName(string rootPath, string input, string targetExtension, bool replace)
        {
            string rawFileName = Path.GetFileNameWithoutExtension(input);

            // support recursive batch conversion 
            string outputSubfolder = "";
            if (Path.GetDirectoryName(input).Length > this.Options.InputPath.Length)
            {
                outputSubfolder = Path.GetDirectoryName(input).Substring(this.Options.InputPath.Length + 1);
            }

            string output = Path.Combine(Path.Combine(rootPath, outputSubfolder), rawFileName + targetExtension);
            
            int num = 0;
            while (!replace && File.Exists(output))
            {
                output = Path.Combine(Path.Combine(rootPath, outputSubfolder), rawFileName + "_" + ++num + targetExtension);
            }
            return output;
        }

        private const int NOT_CONVERTED = 0;
        private const int VALIDATED_AND_OPENED = 1;
        private const int VALIDATED_AND_NOT_OPENED = 2;
        private const int NOT_VALIDATED_AND_OPENED = 3;
        private const int NOT_VALIDATED_AND_NOT_OPENED = 4;
    }

    class ConverterFactory
    {
        private static AbstractConverter wordInstance;
        private static AbstractConverter presentationInstance;
        private static AbstractConverter spreadsheetInstance;
       
        protected ConverterFactory()
        {
        }

        public static AbstractConverter Instance(Direction transformDirection)
        {
            switch (transformDirection)
            {
                case Direction.DocxToOdt:
                case Direction.OdtToDocx:
                    if (wordInstance == null)
                    {
                        wordInstance = new Wordprocessing.Converter();
                    }
                    return wordInstance;
                case Direction.PptxToOdp:
                case Direction.OdpToPptx:
                    if (presentationInstance == null)
                    {
                        presentationInstance = new Sonata.OdfConverter.Presentation.Converter();
                    }
                    return presentationInstance;
                case Direction.XlsxToOds:
                case Direction.OdsToXlsx:
                    if (spreadsheetInstance == null)
                    {
                        spreadsheetInstance = new CleverAge.OdfConverter.Spreadsheet.Converter();
                    }
                    return spreadsheetInstance;
                default:
                    throw new ArgumentException("invalid transform direction type");
            }
        }
    }
}
