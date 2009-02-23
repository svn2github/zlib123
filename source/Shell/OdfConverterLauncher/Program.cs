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
using System.Collections.Generic;
using System.Windows.Forms;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using OdfConverter.OdfConverterLib;
using System.Diagnostics;

namespace OdfConverterLauncher
{
    class OfficeApplication 
    {
        Type _type;
        LateBindingObject _application;
        string _progIdAddin;
        
        public OfficeApplication(string progId, string progIdAddin)
        {
            _progIdAddin = progIdAddin;
            _type = Type.GetTypeFromProgID(progId);
            object instance = null;
            try
            {
                instance = Marshal.GetActiveObject(progId);
            }
            catch (COMException)
            {
                Trace.WriteLine(string.Format("{0} not running yet.", progId));
            }
            
            if (instance == null)
            {
                instance = Activator.CreateInstance(_type);
            }
            _application = new LateBindingObject(instance);
            _application.SetBool("Visible", true);
        }

        public void ImportOdf(string odfFileName)
        {
            try
            {
                // call the exposed add-in method for importing ODF
                _application.Invoke("COMAddIns", _progIdAddin).Invoke("Object").Invoke("importOdfFile", odfFileName);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
                throw;
            }
        }
    }
    
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();

            if (args.Length == 1)
            {
                OfficeApplication app = null;
                
                // args[0] might be a short DOS name but we want the full path name
                FileInfo fi = new FileInfo(args[0]);
                string input = fi.FullName;
                
                try
                {
                    if (input.ToUpper().EndsWith(".ODP") || input.ToUpper().EndsWith(".OTP"))
                    {
                        app = new OfficeApplication("PowerPoint.Application", "OdfPowerPointAddin.Connect");
                    }
                    else if (input.ToUpper().EndsWith(".ODS") || input.ToUpper().EndsWith(".OTS"))
                    {
                        app = new OfficeApplication("Excel.Application", "OdfExcelAddin.Connect");
                    }
                    else if (input.ToUpper().EndsWith(".ODT") || input.ToUpper().EndsWith(".OTT"))
                    {
                        app = new OfficeApplication("Word.Application", "OdfWordAddin.Connect");
                    }
                    app.ImportOdf(input);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.ToString());
                }
            }
        }
    }
}