/* 
 * Copyright (c) 2006 Clever Age, (c) 2008 DIaLOGIKa
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
 *     * Neither the names of copyright holders, nor the names of its contributors
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE 
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
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


namespace OdfConverter.CommandLineTool
{
    /// <summary>
    /// OdfConverter is a CommandLine wrapper for the OpenXML - ODF translators
    /// 
    /// Execute the command without argument to see the options.
    /// </summary>
    public class OdfConverter
    {
        private const string GENERATOR = "OpenXML/ODF Translator Command Line Tool";
        private ConversionReport _report = null;
        private IValidator _ooxValidator = null;
        private IValidator _odfValidator = null;

        // fields to display a progress bar
        private bool _computeSize = false;
        private int _totalSize = 0;
        private int _progress = 0;

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
            try
            {
                ConversionOptions options = OdfConverter.ParseCommandLine(args);
                // complete / correct conversion options
                options.Generator = GENERATOR;
                if (string.IsNullOrEmpty(options.InputFullName))
                {
                    throw new InvalidConversionOptionsException("Input is missing");
                }

                FileInfo fi = new FileInfo(options.InputFullName);
                options.InputFullName = fi.FullName;

                if (!string.IsNullOrEmpty(options.OutputFullName))
                {
                    fi = new FileInfo(options.OutputFullName);
                    options.OutputFullName = fi.FullName;
                }
                if (options.ConversionMode == ConversionMode.Batch)
                {
                    options.InputBaseFolder = options.InputFullName;
                    options.OutputBaseFolder = options.OutputFullName;
                }
                else
                {
                    options.InputBaseFolder = Path.GetDirectoryName(options.InputFullName);
                    options.OutputBaseFolder = Path.GetDirectoryName(options.OutputFullName);
                }

                OdfConverter converter = new OdfConverter();
                converter.proceed(options);
                Console.WriteLine("Done.");
            }
            catch (OdfCommandLineException ex)
            {
                Environment.ExitCode = 1;
                Console.WriteLine("Error when parsing command line: " + ex.Message);
                Console.WriteLine();
                usage();
                return;
            }
            catch (InvalidConversionOptionsException ex)
            {
                Environment.ExitCode = 1;
                Console.WriteLine("Incorrect or missing conversion parameters: " + ex.Message);
                Console.WriteLine();
                usage();
                return;
            }
            catch (Exception ex)
            {
                Environment.ExitCode = 1;
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
                return;
            }
        }

        public OdfConverter()
        {
        }

        private void proceed(ConversionOptions options)
        {
            this._report = new ConversionReport(options.ReportPath, options.LogLevel);
            options.Report = this._report;

            if (options.ConversionMode == ConversionMode.Batch)
            {
                if (File.Exists(options.InputFullName))
                {
                    throw new InvalidConversionOptionsException("Input must be a directory for batch conversions.");
                }
                this.checkBatch(options);
                this.proceedBatch(options);
            }
            else
            {
                if (Directory.Exists(options.InputFullName))
                {
                    throw new InvalidConversionOptionsException("Input must be a file unless one of the batch option is specified.");
                }
                this.checkSingleFile(options);
                this.proceedSingleFile(options);
            }

            options.Report.Close();
        }

        private void proceedBatch(ConversionOptions batchOptions)
        {
            string fileSearchPattern = "";

            switch (batchOptions.ConversionDirection)
            {
                case ConversionDirection.DocxToOdt:
                    fileSearchPattern = "*.docx;*.docm;*.dotx;*.dotm";
                    break;
                case ConversionDirection.XlsxToOds:
                    fileSearchPattern = "*.xlsx;*xlsm;*.xltx;*.xltm";
                    break;
                case ConversionDirection.PptxToOdp:
                    fileSearchPattern = "*.pptx;*pptm;*potx;*potm";
                    break;
                case ConversionDirection.OdtToDocx:
                    fileSearchPattern = "*.odt;*.ott";
                    break;
                case ConversionDirection.OdsToXlsx:
                    fileSearchPattern = "*.ods;*.ots";
                    break;
                case ConversionDirection.OdpToPptx:
                    fileSearchPattern = "*.odp;*.otp";
                    break;
                default:
                    throw new InvalidConversionOptionsException("Unsupported batch mode.");
            }

            string[] files = getFiles(batchOptions.InputBaseFolder, fileSearchPattern, batchOptions.RecursiveMode);

            int nbFiles = files.Length;
            int nbConverted = 0;
            int nbValidated = 0;
            int nbNotValidated = 0;
            this._report.AddComment("Processing " + nbFiles + " file(s).");

            foreach (string input in files)
            {
                // specify new options for each file
                ConversionOptions options = new ConversionOptions();
                options.InputFullName = input;
                options.InputBaseFolder = batchOptions.InputFullName;
                options.OutputBaseFolder = batchOptions.OutputBaseFolder;

                options.AutoDetectDirection = true;
                options.ConversionDirection = batchOptions.ConversionDirection;
                options.ConversionMode = ConversionMode.Batch;
                
                options.Validate = batchOptions.Validate;
                options.ForceOverwrite = batchOptions.ForceOverwrite;

                options.Report = batchOptions.Report;
                options.ReportPath = batchOptions.ReportPath;
                options.LogLevel = batchOptions.LogLevel;

                options.ShowUserInterface = batchOptions.ShowUserInterface;
                options.ShowProgress = batchOptions.ShowProgress;
                
                options.XslPath = batchOptions.XslPath;
                options.SkippedPostProcessors = batchOptions.SkippedPostProcessors;
                options.Packaging = batchOptions.Packaging;

                options.Generator = batchOptions.Generator;

                try
                {
                    this.checkSingleFile(options);

                    int result = this.proceedSingleFile(options);

                    switch (result)
                    {
                        case NOT_CONVERTED:
                            break;
                        case VALIDATED:
                            nbConverted++;
                            nbValidated++;
                            break;
                        case NOT_VALIDATED:
                            nbConverted++;
                            nbNotValidated++;
                            break;
                        default:
                            break;
                    }
                }
                catch (InvalidConversionOptionsException ex)
                {
                    _report.LogInfo(options.InputFullName, "Skipping file {0}. {1}", options.InputFullName, ex.Message);
                }
            }

            this._report.AddComment("Results: " + nbConverted + " of " + nbFiles + " file(s) were converted successfully.");

            if (batchOptions.Validate)
            {
                this._report.AddComment("   " + nbValidated + " file(s) could be validated.");
                this._report.AddComment("   " + nbNotValidated + " file(s) could not be validated.");
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
        private string[] getFiles(string path, string searchPattern, bool recursive)
        {
            string[] m_arExt = searchPattern.Split(';');

            SearchOption option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            List<string> strFiles = new List<string>();
            foreach (string filter in m_arExt)
            {
                strFiles.AddRange(System.IO.Directory.GetFiles(path, filter, option));
            }

            // do not process temp files of Office whose names are starting with "~$"
            List<string> strFilesFiltered = new List<string>();
            foreach (string file in strFiles)
            {
                if (!Path.GetFileName(file).StartsWith("~$"))
                {
                    strFilesFiltered.Add(file);
                }
            }
            return strFilesFiltered.ToArray();
        }

        private int proceedSingleFile(ConversionOptions options)
        {
            bool converted = false;
            bool validated = false;

            _report.LogInfo(options.InputFullName, "Converting {0} into {1}", options.InputFullName.Substring(options.InputBaseFolder.Length + 1), options.OutputFullName.Substring(options.OutputBaseFolder.Length + 1));
            converted = convertFile(options.InputFullName, options.OutputFullName, options);
            if (converted && options.Validate)
            {
                validated = validateFile(options.InputFullName, options.OutputFullName, options.ConversionDirection);
            }
            if (!converted)
            {
                Environment.ExitCode = 1;
                return NOT_CONVERTED;
            }
            else if (validated)
            {
                return VALIDATED;
            }
            else
            {
                return NOT_VALIDATED;
            }
        }

        private bool convertFile(string input, string output, ConversionOptions options)
        {
            try
            {
                DateTime start = DateTime.Now;
                AbstractConverter converter = ConverterFactory.Instance(options.ConversionDirection);
                converter.ExternalResources = options.XslPath;
                converter.SkippedPostProcessors = options.SkippedPostProcessors;
                converter.DirectTransform =
                    options.ConversionDirection == ConversionDirection.OdtToDocx
                    || options.ConversionDirection == ConversionDirection.OdpToPptx
                    || options.ConversionDirection == ConversionDirection.OdsToXlsx;
                converter.Packaging = options.Packaging;

                if (options.ShowProgress)
                {
                    converter.RemoveMessageListeners();
                    converter.AddProgressMessageListener(new AbstractConverter.MessageListener(progressMessageInterceptor));
                    this._computeSize = true;
                    this._progress = 0;
                    this._totalSize = 0;
                    converter.ComputeSize(input);
                    this._computeSize = false;
                }
                
                converter.Transform(input, output, options);

                if (options.ShowProgress)
                {
                    // clear progress bar
                    Console.CursorLeft = 0;
                    Console.Write("                                                                                ");
                    Console.CursorLeft = 0;
                } 
                TimeSpan duration = DateTime.Now - start;
                this._report.LogInfo(input, "Conversion succeeded");
                this._report.LogInfo(input, "Total conversion time: {0}", duration);
                return true;
            }
            catch (EncryptedDocumentException)
            {
                this._report.LogWarning(input, "Conversion failed - Input file is encrypted");
                return false;
            }
            catch (NotAnOdfDocumentException e)
            {
                this._report.LogError(input, "Conversion failed - Input file is not a valid ODF file");
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
            catch (NotAnOoxDocumentException e)
            {
                this._report.LogError(input, "Conversion failed - Input file is not a valid Office OpenXML file");
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
            catch (ZipException e)
            {
                this._report.LogError(input, "Conversion failed - Input file is not a valid file for conversion or might be password protected");
                this._report.LogDebug(input, e.Message);
                return false;
            }
            //Pradeep Nemadi - Bug 1747083 Start
            //IOException is added to fix this bug
            catch (IOException e)
            {
                this._report.LogError(input, "Conversion failed - " + e.Message);
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
            //Pradeep Nemadi - Bug 1747083 end
            catch (Exception e)
            {
                this._report.LogError(input, "Conversion failed - Error during conversion");
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
        }

        /// <summary>
        /// A listener for progress messages from the XSLT.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void progressMessageInterceptor(object sender, EventArgs e)
        {
            if (this._computeSize)
            {
                this._totalSize++;
            }
            else
            {
                this._progress++;
                drawTextProgressBar(this._progress, this._totalSize);                
            }
        }

        /// <summary>
        /// Draw a progress bar at the current cursor position.
        /// Be careful not to Console.WriteLine or anything whilst using this to show progress!
        /// </summary>
        /// <param name="progress">The position of the bar</param>
        /// <param name="total">The amount it counts</param>

        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 66;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 65.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.CursorLeft = position++;
                Console.Write("#");
            }

            //draw unfilled part
            for (int i = position; i < 64; i++)
            {
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 69;
            int percentage = (int)(100.0 * (double)progress / (double)total);
            Console.Write(percentage.ToString() + "%     "); //blanks at the end remove any excess
        }

        private bool validateFile(string input, string output, ConversionDirection transformDirection)
        {
            try
            {
                if (transformDirection == ConversionDirection.OdtToDocx
                    || transformDirection == ConversionDirection.OdsToXlsx
                    || transformDirection == ConversionDirection.OdpToPptx)
                {
                    if (this._ooxValidator == null)
                    {
                        this._report.AddComment("Instantiating OpenXML validator, please wait...");
                        this._ooxValidator = new OoxValidator(this._report);
                        this._report.AddComment("Validator instanciated");
                    }
                    this._ooxValidator.validate(output);
                }
                else
                {
                    if (this._odfValidator == null)
                    {
                        this._report.AddComment("Instantiating ODF validator, please wait...");
                        this._odfValidator = new OdfValidator(this._report);
                        this._report.AddComment("Validator instanciated");
                    }
                    this._odfValidator.validate(output);
                }
                this._report.LogInfo(input, "Converted file {0} is valid", output);
                return true;
            }
            catch (ValidationException e)
            {
                this._report.LogWarning(input, "Converted file {0} is not valid", output);
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
            catch (Exception e)
            {
                this._report.LogError(input, "An unexpected exception occurred when trying to validate {0}", output);
                this._report.LogDebug(input, e.Message + "(" + e.StackTrace + ")");
                return false;
            }
        }

        private static void usage()
        {
            Console.WriteLine("Usage: OdfConverter.exe /I PathOrFilename [/O PathOrFilename] [/<OPTIONS>]");
            Console.WriteLine();
            Console.WriteLine("  Where options are:");
            Console.WriteLine("     /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)");
            Console.WriteLine("     /O PathOrFilename  Name of the output file (or output folder)");
            Console.WriteLine("     /F                 Overwrite existing file(s)");
            Console.WriteLine("     /V                 Validate the result of the transformation against the schemas");
            Console.WriteLine("     /P                 Show conversion progress on the command line");
            Console.WriteLine("     /REPORT Filename   Name of the report file that should be generated (existing files will be replaced)");
            Console.WriteLine("     /LEVEL Level       Level of reporting: 1=DEBUG, 2=INFO, 3=WARNING, 4=ERROR");
            Console.WriteLine(); 
            Console.WriteLine("  Batch options (use one of the /BATCH-<format> options at a time):");
            Console.WriteLine("     /BATCH-ODT         Do a batch conversion over every ODT file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-ODP         Do a batch conversion over every ODP file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-PPTX        Do a batch conversion over every PPTX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-ODS         Do a batch conversion over every ODS file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /BATCH-XLSX        Do a batch conversion over every XLSX file in the input folder (Note: use /F to replace existing files)");
            Console.WriteLine("     /R                 Process subfolders recursively during batch conversion");
            Console.WriteLine(); 
            Console.WriteLine("  Conversion direction options (to disable automatic file type detection):");
            Console.WriteLine("     /ODT2DOCX          Force conversion to DOCX regardless of input file extension");
            Console.WriteLine("     /DOCX2ODT          Force conversion to ODT regardless of input file extension");
            Console.WriteLine("     /ODS2XLSX          Force conversion to XLSX regardless of input file extension");
            Console.WriteLine("     /XLSX2ODS          Force conversion to ODS regardless of input file extension");
            Console.WriteLine("     /ODP2PPTX          Force conversion to PPTX regardless of input file extension");
            Console.WriteLine("     /PPTX2ODP          Force conversion to ODP regardless of input file extension");
            Console.WriteLine(); 
            Console.WriteLine("  Developer options:");
            Console.WriteLine("     /XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)");
            Console.WriteLine("     /NOPACKAGING       Don't package the result of the transformation into a ZIP archive (produce raw XML)");
            Console.WriteLine("     /SKIP Name         Skip a post-processing (provide the post-processor's name)");

        }

        private static ConversionOptions ParseCommandLine(string[] args)
        {
            ConversionOptions options = new ConversionOptions();
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToUpperInvariant().Replace('/', '-').Replace('_', '-'))
                {
                    case "-I":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Input missing");
                        }
                        options.InputFullName = args[i];
                        break;
                    case "-O":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Output missing");
                        }
                        options.OutputFullName = args[i];
                        break;
                    case "-V":
                        options.Validate = true;
                        break;
                    case "-LEVEL":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Level missing");
                        }
                        try
                        {
                            options.LogLevel = (LogLevel)int.Parse(args[i]);
                        }
                        catch (Exception)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1, 2, 3 or 4)");
                        }
                        if ((int)options.LogLevel < 1 || (int)options.LogLevel > 4)
                        {
                            throw new OdfCommandLineException("Wrong level (must be 1, 2, 3 or 4)");
                        }
                        break;
                    case "-R":
                        options.RecursiveMode = true;
                        break;
                    case "-NOPACKAGING":
                        options.Packaging = false;
                        break;
                    case "-BATCH-ODT":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.OdtToDocx;
                        break;
                    case "-BATCH-DOCX":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.DocxToOdt;
                        break;
                    case "-BATCH-ODP":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.OdpToPptx;
                        break;
                    case "-BATCH-PPTX":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.PptxToOdp;
                        break;
                    case "-BATCH-ODS":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.OdsToXlsx;
                        break;
                    case "-BATCH-XLSX":
                        options.ConversionMode = ConversionMode.Batch;
                        options.ConversionDirection = ConversionDirection.XlsxToOds;
                        break;
                    case "-SKIP":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Post processing name missing");
                        }
                        options.SkippedPostProcessors.Add(args[i]);
                        break;
                    case "-REPORT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("Report file missing");
                        }
                        options.ReportPath = args[i];
                        break;
                    case "-XSLT":
                        if (++i == args.Length)
                        {
                            throw new OdfCommandLineException("XSLT path missing");
                        }
                        options.XslPath = args[i];
                        break;
                    case "-DOCX2ODT":
                        options.ConversionDirection = ConversionDirection.DocxToOdt;
                        options.AutoDetectDirection = false;
                        break;
                    case "-ODT2DOCX":
                        options.ConversionDirection = ConversionDirection.OdtToDocx;
                        options.AutoDetectDirection = false;
                        break;
                    case "-XLSX2ODS":
                        options.ConversionDirection = ConversionDirection.XlsxToOds;
                        options.AutoDetectDirection = false;
                        break;
                    case "-ODS2XLSX":
                        options.ConversionDirection = ConversionDirection.OdsToXlsx;
                        options.AutoDetectDirection = false;
                        break;
                    case "-PPTX2ODP":
                        options.ConversionDirection = ConversionDirection.PptxToOdp;
                        options.AutoDetectDirection = false;
                        break;
                    case "-ODP2PPTX":
                        options.ConversionDirection = ConversionDirection.OdpToPptx;
                        options.AutoDetectDirection = false;
                        break;
                    case "-F":
                        options.ForceOverwrite = true;
                        break;
                    case "-P":
                        options.ShowProgress = true;
                        break;
                    default:
                        if (args[i].ToUpperInvariant().Replace('/', '-').Replace('_', '-').StartsWith("-"))
                        {
                            throw new OdfCommandLineException(string.Format("Invalid parameter {0} found.", args[i]));
                        }
                        else if (string.IsNullOrEmpty(options.InputFullName))
                        {
                            options.InputFullName = args[i];
                        }
                        break;
                }
            }
            return options;
        }

        private void checkBatch(ConversionOptions options)
        {
            if (!Directory.Exists(options.InputFullName))
            {
                throw new InvalidConversionOptionsException(string.Format("Input folder {0} does not exist", options.InputFullName));
            }
            if (File.Exists(options.OutputFullName))
            {
                throw new InvalidConversionOptionsException(string.Format("Output {0} must be a folder", options.OutputFullName));
            }
            if (string.IsNullOrEmpty(options.OutputFullName))
            {
                // use input folder as output folder
                options.OutputFullName = options.InputFullName;
            }
            if (!Directory.Exists(options.OutputFullName))
            {
                try
                {
                    Directory.CreateDirectory(options.OutputFullName);
                }
                catch (Exception)
                {
                    throw new InvalidConversionOptionsException(string.Format("Cannot create output folder {0}.", options.OutputFullName));
                }
            }
        }

        private void checkSingleFile(ConversionOptions options)
        {
            if (!File.Exists(options.InputFullName))
            {
                throw new InvalidConversionOptionsException(string.Format("Input file {0} does not exist.", options.InputFullName));
            }

            string inputExtension = Path.GetExtension(options.InputFullName).ToLowerInvariant();

            if (options.AutoDetectDirection)
            {
                // auto-detect transform direction based on file extension
                switch (inputExtension)
                {
                    case ".odt":
                    case ".ott":
                        options.ConversionDirection = ConversionDirection.OdtToDocx;
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        options.ConversionDirection = ConversionDirection.DocxToOdt;
                        break;
                    case ".odp":
                    case ".otp":
                        options.ConversionDirection = ConversionDirection.OdpToPptx;
                        break;
                    case ".pptx":
                    case ".pptm":
                    case ".potx":
                    case ".potm":
                        options.ConversionDirection = ConversionDirection.PptxToOdp;
                        break;
                    case ".ods":
                    case ".ots":
                        options.ConversionDirection = ConversionDirection.OdsToXlsx;
                        break;
                    case ".xlsx":
                    case ".xlsm":
                    case ".xltx":
                    case ".xltm":
                        options.ConversionDirection = ConversionDirection.XlsxToOds;
                        break;
                    default:
                        throw new InvalidConversionOptionsException("Input file extension [" + inputExtension + "] is not supported.");
                }
            }

            // detect document type
            switch (inputExtension)
            {
                case ".dotx":
                case ".dotm":
                case ".xltx":
                case ".xltm":
                case ".potx":
                case ".potm":
                case ".ott":
                case ".ots":
                case ".otp":
                    options.DocumentType = DocumentType.Template;
                    break;
                default:
                    options.DocumentType = DocumentType.Document;
                    break;
            }

            if (string.IsNullOrEmpty(options.OutputFullName) || options.ConversionMode == ConversionMode.Batch)
            {
                string outputExtension = getTargetExtension(options);
                string outputPath = options.OutputBaseFolder;
                if (string.IsNullOrEmpty(options.OutputBaseFolder))
                {
                    // we take input path
                    options.OutputBaseFolder = options.InputBaseFolder;
                }
                options.OutputFullName = generateOutputName(options.OutputBaseFolder, options.InputFullName, outputExtension, options);
            }
            else if (options.ConversionMode == ConversionMode.SingleDocument)
            {
                if (File.Exists(options.OutputFullName) && !options.ForceOverwrite)
                {
                    throw new InvalidConversionOptionsException("The specified output file already exists. Use /F to overwrite existing files.");
                }
            }
        }

        private string getTargetExtension(ConversionOptions options)
        {
            string targetExtension = "";
            if (!options.Packaging)
            {
                targetExtension = ".xml";
            }
            else if (!string.IsNullOrEmpty(options.OutputFullName))
            {
                targetExtension = Path.GetExtension(options.OutputFullName);
            }
            else
            {
                switch (Path.GetExtension(options.InputFullName).ToLowerInvariant())
                {
                    case ".docx":
                    case ".docm":
                        targetExtension = ".odt";
                        break;
                    case ".dotx":
                    case ".dotm":
                        targetExtension = ".ott";
                        break;
                    case ".xlsx":
                    case ".xlsm":
                        targetExtension = ".ods";
                        break;
                    case ".xltx":
                    case ".xltm":
                        targetExtension = ".ots";
                        break;
                    case ".pptx":
                    case ".pptm":
                        targetExtension = ".odp";
                        break;
                    case ".potx":
                    case ".potm":
                        targetExtension = ".otp";
                        break;
                    case ".odt":
                        targetExtension = ".docx";
                        break;
                    case ".ott":
                        targetExtension = ".dotx";
                        break;
                    case ".ods":
                        targetExtension = ".xlsx";
                        break;
                    case ".ots":
                        targetExtension = ".xltx";
                        break;
                    case ".odp":
                        targetExtension = ".pptx";
                        break;
                    case ".otp":
                        targetExtension = ".potx";
                        break;
                    default:
                        // unknown file extension
                        switch (options.ConversionDirection)
                        {
                            case ConversionDirection.DocxToOdt:
                                targetExtension = ".odt";
                                break;
                            case ConversionDirection.OdtToDocx:
                                targetExtension = ".docx";
                                break;
                            case ConversionDirection.PptxToOdp:
                                targetExtension = ".odp";
                                break;
                            case ConversionDirection.OdpToPptx:
                                targetExtension = ".pptx";
                                break;
                            case ConversionDirection.OdsToXlsx:
                                targetExtension = ".xlsx";
                                break;
                            case ConversionDirection.XlsxToOds:
                                targetExtension = ".ods";
                                break;
                            default:
                                throw new InvalidConversionOptionsException(string.Format("Unable to determine output file extension for input file {0}.", options.InputFullName));
                        }
                        break;
                }
            }
            return targetExtension;
        }

        private string generateOutputName(string rootPath, string input, string targetExtension, ConversionOptions options)
        {
            string rawFileName = Path.GetFileNameWithoutExtension(input);

            // support recursive batch conversion 
            string outputSubfolder = "";
            if (Path.GetDirectoryName(input).Length > options.InputBaseFolder.Length)
            {
                outputSubfolder = Path.GetDirectoryName(input).Substring(options.InputBaseFolder.Length + 1);
            }

            string output = Path.Combine(Path.Combine(rootPath, outputSubfolder), rawFileName + targetExtension);

            int num = 0;
            while (!options.ForceOverwrite && File.Exists(output))
            {
                output = Path.Combine(Path.Combine(rootPath, outputSubfolder), rawFileName + "_" + ++num + targetExtension);
            }
            return output;
        }

        private const int NOT_CONVERTED = 0;
        private const int VALIDATED = 1;
        private const int NOT_VALIDATED = 2;
    }

    class ConverterFactory
    {
        private static AbstractConverter wordInstance;
        private static AbstractConverter presentationInstance;
        private static AbstractConverter spreadsheetInstance;

        protected ConverterFactory()
        {
        }

        public static AbstractConverter Instance(ConversionDirection transformDirection)
        {
            switch (transformDirection)
            {
                case ConversionDirection.DocxToOdt:
                case ConversionDirection.OdtToDocx:
                    if (wordInstance == null)
                    {
                        wordInstance = new Wordprocessing.Converter();
                    }
                    return wordInstance;
                case ConversionDirection.PptxToOdp:
                case ConversionDirection.OdpToPptx:
                    if (presentationInstance == null)
                    {
                        presentationInstance = new Sonata.OdfConverter.Presentation.Converter();
                    }
                    return presentationInstance;
                case ConversionDirection.XlsxToOds:
                case ConversionDirection.OdsToXlsx:
                    if (spreadsheetInstance == null)
                    {
                        spreadsheetInstance = new CleverAge.OdfConverter.Spreadsheet.Converter();
                    }
                    return spreadsheetInstance;
                default:
                    throw new ArgumentException("Invalid transform direction type");
            }
        }
    }
}
