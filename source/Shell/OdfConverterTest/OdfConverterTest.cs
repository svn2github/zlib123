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

namespace CleverAge.OdfConverter.CommandLineTool
{
    enum ControlType : int
    {
        CTRL_C_EVENT = 0,
        CTRL_BREAK_EVENT = 1,
        CTRL_CLOSE_EVENT = 2,
        CTRL_LOGOFF_EVENT = 5,
        CTRL_SHUTDOWN_EVENT = 6
    };

    delegate int ControlHandlerFonction(ControlType control);

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

    /// <summary>
    /// ODFConverterTest is a CommandLine Program to test the conversion
    /// of an OpenDocument file into an OpenXML file.
    /// 
    /// Execute the command without argument to see the options.
    /// </summary>
    public class OdfConverterTest
    {
        private string input = null;                     // input path
        private string output = null;                    // output path
        private bool batchOdt = false;                   // do batch transform on ODT files
        private bool batchDocx = false;                  // do batch transform on DOCX files
        private bool validate = false;                   // validate the result of the transformations
        private bool open = false;                       // try to open the result of the transformations
        private bool recursiveMode = false;              // go in subfolders ?
        private bool replace = false;					 // override existing files ?
        private string reportPath = null;                // file to save report
        private int reportLevel = Report.INFO_LEVEL;     // file to save report
        private string xslPath = null;                   // Path to an external stylesheet
        private ArrayList skipedPostProcessors = null;   // Post processors to skip (identified by their names)
        private bool packaging = true;                   // Build the zip archive after conversion

        private enum Direction { OdtToDocx, DocxToOdt };
	   	private Direction transformDirection = Direction.OdtToDocx; // direction of conversion
		private bool transformDirectionOverride = false; // whether conversion direction has been specified
        private Report report = null;
        private Word word = null;
        private OoxValidator ooxValidator = null;
        private OdfValidator odfValidator = null;

#if MONO
		static bool SetConsoleCtrlHandler(ControlHandlerFonction handlerRoutine, bool add) 
		{ return true; }
#else
		[DllImport("kernel32")]
        static extern bool SetConsoleCtrlHandler(ControlHandlerFonction handlerRoutine, bool add);
#endif

        int MyHandler(ControlType control)
        {
            //Console.WriteLine("MyHandler: " + control.ToString());
            if (word != null)
            {
                try
                {
                    word.Quit();
                }
                catch
                {
                    Console.WriteLine("Unable to close Word, please close it manually");
                }
            }
            return 0;
        }

        /// <summary>
        /// Main program.
        /// </summary>
        /// <param name="args">Command Line arguments</param>
        public static void Main(String[] args)
        {
            OdfConverterTest tester = new OdfConverterTest();
            ControlHandlerFonction myHandler = new ControlHandlerFonction(tester.MyHandler);
            SetConsoleCtrlHandler(myHandler, true);
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
                tester.Proceed();
                Console.WriteLine("Done.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception raised when running test : " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private OdfConverterTest()
        {
            this.skipedPostProcessors = new ArrayList();
        }

        private void Proceed()
        {
            this.report = new Report(this.reportPath, this.reportLevel);
            if (this.batchOdt)
            {
                this.ProceedBatchOdt();
            }
            else if (this.batchDocx)
            {
                this.ProceedBatchDocx();
            }
            else
            {

                // instanciate word if needed
               	if (this.transformDirection == Direction.OdtToDocx && this.open)
                {
                    word = new Word();
                    word.Visible = false;
                }

 				this.ProceedSingleFile(this.input, this.output, this.transformDirection);
                
                // close word if needed
                if (this.open)
                {
                    word.Quit();
                }
            }
            this.report.Close();
        }

        private void ProceedBatchOdt()
        {
            // instanciate word if needed
            if (this.open)
            {
                this.word = new Word();
                this.word.Visible = false;
            }

            // instanciate validator if needed
            if (this.validate)
            {
                this.report.AddComment("Instanciating validator, please wait...");
                this.ooxValidator = new OoxValidator(this.report);
                this.report.AddComment("Validator instanciated");
            }

            SearchOption option = SearchOption.TopDirectoryOnly;
            if (this.recursiveMode)
            {
                option = SearchOption.AllDirectories;
            }
            string[] files = Directory.GetFiles(this.input, "*.odt", option);
            int nbFiles = files.Length;
            int nbConverted = 0;
            int nbValidatedAndOpened = 0;
            int nbValidatedAndNotOpened = 0;
            int nbNotValidatedAndOpened = 0;
            int nbNotValidatedAndNotOpened = 0;
            this.report.AddComment("Processing " + nbFiles + " ODT file(s)");
            foreach (string input in files)
            {
                string output = this.GenerateOutputName(this.output, input, ".docx", this.replace);
                int result = this.ProceedSingleFile(input, output, Direction.OdtToDocx);
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
            if (this.open)
            {
                word.Quit();
            }
            this.report.AddComment("Results: " + nbConverted + " file(s) over " + nbFiles + " were converted, among them:");
            this.report.AddComment("   " + nbValidatedAndOpened + " were validated and sucessfully opened in Word");
            this.report.AddComment("   " + nbValidatedAndNotOpened + " were validated but could not be opened in Word");
            this.report.AddComment("   " + nbNotValidatedAndOpened + " were not validated but were sucessfully opened in Word");
            this.report.AddComment("   " + nbNotValidatedAndNotOpened + " were not validated and could not be opened in Word");
        }

        private void ProceedBatchDocx()
        {
            // instanciate validator if needed
            if (this.validate)
            {
                this.report.AddComment("Instanciating validator, please wait...");
                this.odfValidator = new OdfValidator(this.report);
                this.report.AddComment("Validator instanciated");
            }
            SearchOption option = SearchOption.TopDirectoryOnly;
            if (this.recursiveMode)
            {
                option = SearchOption.AllDirectories;
            }
            string[] files = Directory.GetFiles(this.input, "*.docx", option);
            foreach (string input in files)
            {
                string output = this.GenerateOutputName(this.output, input, ".odt", this.replace);
                this.ProceedSingleFile(input, output, Direction.DocxToOdt);
            }
        }

        private int ProceedSingleFile(string input, string output, Direction transformDirection)
        {
            bool converted = false;
            bool validated = false;
            bool opened = false;
            report.AddLog(input, "Converting file: " + input + " into " + output, Report.INFO_LEVEL);
            converted = ConvertFile(input, output, transformDirection);
            if (converted && this.validate)
            {
                validated = ValidateFile(input, output, transformDirection);
            }
            if (converted && this.open)
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
                Converter converter = new Converter();
                converter.ExternalResources = this.xslPath;
                converter.SkipedPostProcessors = this.skipedPostProcessors;
                converter.DirectTransform = transformDirection == Direction.OdtToDocx;
                converter.Packaging = this.packaging;
                converter.Transform(input, output);
                TimeSpan duration = DateTime.Now - start;
                this.report.AddLog(input, "Conversion succeeded", Report.INFO_LEVEL);
                this.report.AddLog(input, "Total conversion time: " + duration, Report.INFO_LEVEL);
                return true;
            }
            catch (EncryptedDocumentException)
            {
                this.report.AddLog(input, "Conversion failed - Input file is encrypted", Report.WARNING_LEVEL);
                return false;
            }
            catch (NotAnOdfDocumentException e)
            {
                this.report.AddLog(input, "Conversion failed - Input file is not a valid ODF file", Report.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }
            catch (NotAnOoxDocumentException e)
            {
                this.report.AddLog(input, "Conversion failed - Input file is not a valid DOCX file", Report.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }
            catch (ZipException e)
            {
                this.report.AddLog(input, "Conversion failed - Error during conversion (ZipException)", Report.ERROR_LEVEL);
                this.report.AddLog(input, e.Message, Report.DEBUG_LEVEL);
                return false;
            }
            catch (Exception e)
            {
                this.report.AddLog(input, "Conversion failed - Error during conversion", Report.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }
        }

        private bool ValidateFile(string input, string output, Direction transformDirection)
        {
            if (transformDirection == Direction.OdtToDocx)
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
                    this.report.AddLog(input, "Converted file (" + output + ") is valid", Report.INFO_LEVEL);
                    return true;
                }
                catch (OoxValidatorException e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") is invalid", Report.WARNING_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                    return false;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "An unexpected exception occured when trying to validate " + output, Report.ERROR_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
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
                    this.report.AddLog(input, "Converted file (" + output + ") is valid", Report.INFO_LEVEL);
                    return true;
                }
                catch (OdfValidatorException e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") is invalid", Report.WARNING_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                    return false;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "An unexpected exception occured when trying to validate " + output, Report.ERROR_LEVEL);
                    this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
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
                    this.report.AddLog(input, "Converted file opened successfully in Word", Report.INFO_LEVEL);
                    return true;
                }
                catch (Exception e)
                {
                    this.report.AddLog(input, "Converted file (" + output + ") could not be opened in Word", Report.ERROR_LEVEL);
                    this.report.AddLog(input, e.GetType().Name + ": " + e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                    return false;
                }
            }
            return false;
        }

        private static void usage()
        {
            Console.WriteLine("Usage: OdfConverterTest.exe /I PathOrFilename [/O PathOrFilename] [/BATCH-ODT] [/BATCH-DOCX] [/V] [/OPEN] [/XSLT Path] [/NOPACKAGING] [/SKIP name] [/REPORT Filename] [/LEVEL Level] ");
            Console.WriteLine("  Where options are:");
            Console.WriteLine("     /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)");
            Console.WriteLine("     /O PathOrFilename  Name of the output file (or output folder)");
            Console.WriteLine("     /F                 Replace existing file");
            Console.WriteLine("     /BATCH-ODT         Do a batch conversion over every ODT file in the input folder (note: existing files will be replaced)");
            Console.WriteLine("     /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (note: existing files will be replaced)");
            Console.WriteLine("     /V                 Validate the result of the transformation against the schemas");
            Console.WriteLine("     /OPEN              Try to open the converted files (works only for ODF->OOX, Microsoft Word required)");
            Console.WriteLine("     /XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)");
            Console.WriteLine("     /NOPACKAGING       Don't package the result of the transformation into a ZIP archive (produce raw XML)");
            Console.WriteLine("     /SKIP name         Skip a post-processing (provide the post-processor's name)");
            Console.WriteLine("     /REPORT Filename   Name of the report file that must be generated (existing files will be replaced)");
            Console.WriteLine("     /LEVEL Level       Level of reporting: 1=DEBUG, 2=INFO, 3=WARNING, 4=ERROR");
          	Console.WriteLine("     /ODT2DOCX          Force conversion to DOCX regardless of input file extension");
			Console.WriteLine("     /DOCX2ODT          Force conversion to ODT regardless of input file extension");
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
                switch (args[i])
                {
                    case "/I":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Input missing");
                        }
                        this.input = args[i];
                        break;
                    case "/O":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Output missing");
                        }
                        this.output = args[i];
                        break;
                    case "/V":
                        this.validate = true;
                        break;
                    case "/OPEN":
                        this.open = true;
                        break;
                    case "/LEVEL":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Level missing");
                        }
                        try
                        {
                            this.reportLevel = int.Parse(args[i]);
                        }
                        catch (Exception)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1,2 3 or 4)");
                        }
                        if (this.reportLevel < 1 || this.reportLevel > 4)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1,2 3 or 4)");
                        }
                        break;
                    case "/R":
                        this.recursiveMode = true;
                        break;
                    case "/NOPACKAGING":
                        this.packaging = false;
                        break;
                    case "/BATCH-ODT":
                        this.batchOdt = true;
                        break;
                    case "/BATCH-DOCX":
                        this.batchDocx = true;
                        break;
                    case "/SKIP":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Post processing name missing");
                        }
                        this.skipedPostProcessors.Add(args[i]);
                        break;
                    case "/REPORT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Report file missing");
                        }
                        this.reportPath = args[i];
                        break;
                    case "/XSLT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("XSLT path missing");
                        }
                        this.xslPath = args[i];
                        break;
                    case "/CV":
                        OoxValidator.test();
                        break;
                   	case "/DOCX2ODT":
						this.transformDirection = Direction.DocxToOdt;
						this.transformDirectionOverride = true;
						// System.Console.WriteLine("Override to odt\n");
						break;
					case "/ODT2DOCX":
						this.transformDirection = Direction.OdtToDocx;
						this.transformDirectionOverride = true;
						// System.Console.WriteLine("Override to docx\n");
						break;
					case "/F":
						this.replace = true;
						break;
                    default:
                        break;
                }
            }
        }

        private void CheckPaths()
        {
            if (this.input == null)
            {
                throw new OdfCommandLineException("Input is missing");
            }
            if (this.batchDocx || this.batchOdt)
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
            if (!Directory.Exists(this.input))
            {
                throw new OdfCommandLineException("Input folder does not exist");
            }
            if (File.Exists(this.output))
            {
                throw new OdfCommandLineException("Output must be a folder");
            }
            if (this.output == null || this.output.Length == 0)
            {
                // use input folder as output folder
                this.output = this.input;
            }
            if (!Directory.Exists(this.output))
            {
                try
                {
                    Directory.CreateDirectory(this.output);
                }
                catch (Exception)
                {
                    throw new OdfCommandLineException("Cannot create output folder");
                }
            }
        }

        private void CheckSingleFile()
        {
            if (!File.Exists(this.input))
            {
                throw new OdfCommandLineException("Input file does not exist");
            }

            string extension = null;
            if (transformDirectionOverride)
            {
            	extension = transformDirection == Direction.OdtToDocx ? ".odt" : ".docx";
            }
            else
            {
                extension = ".docx";
                if (extension.Equals(Path.GetExtension(this.input).ToLowerInvariant()))
                {
                	this.transformDirection = Direction.DocxToOdt;
                	extension = ".odt";
                }
            }
            if (!this.packaging)
            {
                extension = ".xml";
            }

            if (!transformDirectionOverride && 
                !File.Exists(this.output) && (this.output == null || !extension.Equals(Path.GetExtension(this.output).ToLowerInvariant())))
            {
                string outputPath = this.output;
                if (outputPath == null)
                {
                    // we take input path
                    outputPath = Path.GetDirectoryName(this.input);
                }
                this.output = GenerateOutputName(outputPath, this.input, extension, this.replace);
            }
        }

        private string GenerateOutputName(string rootPath, string input, string extension, bool replace)
        {
            string rawFileName = Path.GetFileNameWithoutExtension(input);
            string output = Path.Combine(rootPath, rawFileName + extension);
            int num = 0;
            while (!replace && File.Exists(output))
            {
                output = Path.Combine(rootPath, rawFileName + "_" + ++num + extension);
            }
            return output;
        }

        private const int NOT_CONVERTED = 0;
        private const int VALIDATED_AND_OPENED = 1;
        private const int VALIDATED_AND_NOT_OPENED = 2;
        private const int NOT_VALIDATED_AND_OPENED = 3;
        private const int NOT_VALIDATED_AND_NOT_OPENED = 4;
    }

    public class Report
    {
        public const int DEBUG_LEVEL = 1;
        public const int INFO_LEVEL = 2;
        public const int WARNING_LEVEL = 3;
        public const int ERROR_LEVEL = 4;

        private StreamWriter writer = null;
        private int level = INFO_LEVEL;

        public Report(string filename, int level)
        {
            this.level = level;
            if (filename != null)
            {
                this.writer = new StreamWriter(new FileStream(filename, FileMode.Create, FileAccess.Write));
                Console.WriteLine("Using report file: " + filename);
            }
        }

        public void AddComment(string message)
        {
            string text = "*** " + message;

            if (this.writer != null)
            {
                this.writer.WriteLine(text);
                this.writer.Flush();
            }
            Console.WriteLine(text);
        }

        public void AddLog(string filename, string message, int level)
        {
            if (level >= this.level)
            {
                string label = null;
                switch (level)
                {
                    case 4:
                        label = "ERROR";
                        break;
                    case 3:
                        label = "WARNING";
                        break;
                    case 2:
                        label = "INFO";
                        break;
                    default:
                        label = "DEBUG";
                        break;
                }
                string text = "[" + label + "]" + "[" + filename + "] " + message;

                if (this.writer != null)
                {
                    this.writer.WriteLine(text);
                    this.writer.Flush();
                }
                Console.WriteLine(text);
            }
        }

        public void Close()
        {
            if (this.writer != null)
            {
                this.writer.Close();
                this.writer = null;
            }
        }

    }
}
