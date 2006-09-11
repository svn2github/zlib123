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
using CleverAge.OdfConverter.OdfConverterLib;
using System.Diagnostics;
using CleverAge.OdfConverter.OdfZipUtils;

namespace CleverAge.OdfConverter.OdfConverterTest
{
    /// <summary>
    /// ODFConverterTest is a CommandLine Program to test the conversion
    /// of an OpenDocument file into an OpenXML file.
    /// 
    /// Execute the command without argument to see the options.
    /// </summary>
    public class OdfConverterTest
    {
        private string input = null;           // input path
        private string output = null;          // output path
        private bool validate = false;         // validate the result of the transformations
        private bool debug = false;            // debug mode
        private bool recursiveMode = false;    // go in subfolders ?
        private string reportPath = null;      // file to save report
        private string xslPath = null;         // Path to an external stylesheet
        private bool displayDuration = false;  // display time taken by conversion
        private bool isDirectTransform = true; // true if the transfo is from ODF to OOX

        private OdfConverterTest(String input, String output, bool validate, bool debug, bool recursiveMode, String reportPath, String xslPath, bool displayDuration, bool isDirectTransform)
        {
            this.input = input;
            this.output = output;
            this.validate = validate;
            this.debug = debug;
            this.recursiveMode = recursiveMode;
            this.reportPath = reportPath;
            this.xslPath = xslPath;
            this.displayDuration = displayDuration;
            this.isDirectTransform = isDirectTransform;
        }

        /// <summary>
        /// Main program.
        /// </summary>
        /// <param name="args">Command Line arguments</param>
        public static void Main(String[] args)
        {
            string input = null;
            string output = null;
            bool validate = false;
            bool debug = false;
            bool recursiveMode = false;
            string reportPath = null;
            string xslPath = null;
            bool doReturn = false;
            bool displayDuration = false;
            bool isDirectTransform = true;

            //Debug.Listeners.Add(new TextWriterTraceListener(Console.Out));
            //Debug.AutoFlush = true;
            //Debug.Indent();


            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "/I":
                        if (++i == args.Length)
                        {
                            usage();
                            return;
                        }
                        input = args[i];
                        break;
                    case "/O":
                        if (++i == args.Length)
                        {
                            usage();
                            return;
                        }
                        output = args[i];
                        break;
                    case "/V":
                        validate = true;
                        break;
                    case "/T":
                        displayDuration = true;
                        break;
                    case "/DEBUG":
                        debug = true;
                        break;
                    case "/R":
                        recursiveMode = true;
                        break;
                    case "/REPORT":
                        if (++i == args.Length)
                        {
                            usage();
                            return;
                        }
                        reportPath = args[i];
                        break;
                    case "/XSLT":
                        if (++i == args.Length)
                        {
                            usage();
                            return;
                        }
                        xslPath = args[i];
                        break;
                    case "/CV":
                        OoxValidator.test();
                        doReturn = true;
                        break;
                    default:
                        break;
                }
            }
            if (doReturn) { return; }
            if (input == null)
            {
                usage();
                return;
            }
            if (input.ToLowerInvariant().EndsWith(".docx"))
            {
                isDirectTransform = false;
            }
            string extension = ".docx";
            if (!isDirectTransform)
            {
                extension = ".odt";
            }
            if (output == null || Directory.Exists(output))
            {
                // generate output filename
                string rootPath = "";
                if (output != null)
                {
                    rootPath = output;
                }
                int index = input.LastIndexOf(".");
                string root;
                if (index > 0)
                {
                    root = input.Substring(0, index);
                }
                else
                {
                    root = input;
                }
                try
                {
                    output = Path.Combine(rootPath, root + extension);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine("Error: invalid filename provided as input");
                    return;
                }
                int i = 1;
                while (File.Exists(output) || Directory.Exists(output))
                {
                    output = Path.Combine(rootPath, root + "_" + i + extension);
                    i++;
                }
            }
            else
            {
                string current_extension = null;
                try
                {
                    current_extension = Path.GetExtension(output);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine("Error: invalid path provided as output");
                    return;
                }
                if (String.IsNullOrEmpty(current_extension))
                {
                    // add extension
                    output += extension;
                }
            }
            new OdfConverterTest(input, output, validate, debug, recursiveMode, reportPath, xslPath, displayDuration, isDirectTransform).run();
        }

        private void run()
        {
            Console.WriteLine("Running with " + this.input + " > " + this.output);

            // transformation
            try
            {
                DateTime start = DateTime.Now;
                if (xslPath == null)
                {
                    if (this.isDirectTransform)
                    {
                        new Converter().OdfToOox(this.input, this.output);
                    }
                    else
                    {
                        new Converter().OoxToOdf(this.input, this.output);
                    }
                }
                else
                {
                    if (this.isDirectTransform)
                    {
                        new Converter().OdfToOoxWithExternalResources(this.input, this.output, this.xslPath);
                    }
                    else
                    {
                        new Converter().OoxToOdfWithExternalResources(this.input, this.output, this.xslPath);
                    }
                }
                TimeSpan duration = DateTime.Now - start;

                // validation
                if (this.validate && this.isDirectTransform)
                {
                    try
                    {
                        new OoxValidator().validate(this.output);
                    }
                    catch (OoxValidatorException e)
                    {
                        Console.WriteLine("The file is not valid: " + e.Message);
                        if (File.Exists(this.output + ".NOT_VALID"))
                        {
                            File.Delete(this.output + ".NOT_VALID");
                        }
                        File.Move(this.output, this.output + ".NOT_VALID");
                    }
                }
                if (this.displayDuration)
                {
                    Console.WriteLine("Total conversion time: " + duration);
                }
            }
            catch (NotAnOdfDocumentException e)
            {
                Console.WriteLine("Error: " + input + " is not a valid ODF file");

            }
            catch (NotAnOoxDocumentException e)
            {
                Console.WriteLine("Error: " + input + " is not a valid DOCX file");

            }
            catch (ZipException e)
            {
                Console.WriteLine(e.Message);
                if (debug)
                {
                    Console.WriteLine(e.StackTrace);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occured during the process (" + e.Message + ")");
                if (debug)
                {
                    Console.WriteLine(e.StackTrace);
                }
            }
        }

        private static void usage()
        {
            Console.WriteLine("Usage: OdfConverterTest.exe /I Filename [/O PathOrFilename] [/V] [/XSLT Path]");
            Console.WriteLine("  Where:");
            Console.WriteLine("     /I Filename        Name of the ODF file to transform");
            Console.WriteLine("     /O PathOrFilename  Path of the folder where to put the output file (must exist) or name of the output file");
            Console.WriteLine("     /V                 Validate the result of the transformation against OpenXML schemas");
            Console.WriteLine("     /T                 Display conversion time");
            Console.WriteLine("     /DEBUG             Debug mode");
            Console.WriteLine("     /XSLT Path         Path to a folder containing XSLT files (must be the same as used in the lib)");
        }
    }
}
