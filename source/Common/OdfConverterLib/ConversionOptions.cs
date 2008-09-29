/*
 * Copyright (c) 2008 DIaLOGIKa
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
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    public enum Direction
    {
        None,
        OdtToDocx,
        DocxToOdt,
        OdsToXlsx,
        XlsxToOds,
        OdpToPptx,
        PptxToOdp
    };

    public class ConversionOptions
    {
        private string  _inputPath = null;                  // input path
        private string  _outputPath = null;                 // output path
        private bool    _validate = false;                  // validate the result of the transformations
        private bool    _open = false;                      // try to open the result of the transformations
        private bool    _recursiveMode = false;             // go in subfolders ?
        private bool    _forceOverwrite = false;			// override existing files ?
        private string  _reportPath = null;                 // file to save report
        private int     _reportLevel = ConversionReport.INFO_LEVEL;   // file to save report
        private string  _xslPath = null;                    // Path to an external stylesheet
        private List<string> _skippedPostProcessors = new List<string>();  // Post processors to skip (identified by their names)
        private bool    _packaging = true;                  // Build the zip archive after conversion
        private Direction _transformDirection = Direction.OdtToDocx; // direction of conversion

        private bool _showUserInterface = false;
        private string _generator;  

        /// <summary>
        /// Determines whether a progress dialog and other dialogs will be shown. Default is <code>false</code>.
        /// </summary>
        public bool ShowUserInterface
        {
            get { return _showUserInterface; }
            set { _showUserInterface = value; }
        }
       
        /// <summary>
        /// This property contains information about the converter version and environment
        /// It will be written to the document's meta data
        /// </summary>
        public string Generator
        {
            get { return _generator; }
            set { _generator = value; }
        }

        public string InputPath
        {
            get { return _inputPath; }
            set { _inputPath = value; }
        }

        public string OutputPath
        {
            get { return _outputPath; }
            set { _outputPath = value; }
        }

        public bool Validate
        {
            get { return _validate; }
            set { _validate = value; }
        }

        public bool Open
        {
            get { return _open; }
            set { _open = value; }
        }

        public bool RecursiveMode
        {
            get { return _recursiveMode; }
            set { _recursiveMode = value; }
        }

        public bool ForceOverwrite
        {
            get { return _forceOverwrite; }
            set { _forceOverwrite = value; }
        }

        public string ReportPath
        {
            get { return _reportPath; }
            set { _reportPath = value; }
        }

        public int ReportLevel
        {
            get { return _reportLevel; }
            set { _reportLevel = value; }
        }

        public string XslPath
        {
            get { return _xslPath; }
            set { _xslPath = value; }
        }

        public List<string> SkippedPostProcessors
        {
            get { return _skippedPostProcessors; }
            set { _skippedPostProcessors = value; }
        }

        public bool Packaging
        {
            get { return _packaging; }
            set { _packaging = value; }
        }

        
        public Direction TransformDirection
        {
            get { return _transformDirection; }
            set { _transformDirection = value; }
        }
         
        //public _OutputFileType
    }
}
 