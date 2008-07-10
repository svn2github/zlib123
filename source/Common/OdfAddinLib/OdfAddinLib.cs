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
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using CleverAge.OdfConverter.OdfConverterLib;
using System.Runtime.InteropServices;
using OdfConverter.OdfConverterLib;

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    //  Chained resource managers
    /// </summary>
    internal class ChainResourceManager : System.Resources.ResourceManager
    {
        private ArrayList managers;

        public ChainResourceManager()
        {
            this.managers = new ArrayList();
        }

        public void Add(System.Resources.ResourceManager manager)
        {
            managers.Add(manager);
        }

        public override string GetString(string key)
        {
            for (int i = this.managers.Count - 1; i >= 0; i--)
            {
                System.Resources.ResourceManager manager = (System.Resources.ResourceManager)this.managers[i];
                if (manager != null)
                {
                    string value = manager.GetString(key);
                    if (value != null && value.Length > 0)
                    {
                        return value;
                    }
                }
            }
            return null;
        }
    }

    /// <summary>
    /// Base class MS Office add-in implementations.
    /// </summary>
    //[ComVisible(true)]
    public class OdfAddinLib : IOdfConverter
    {
        private AbstractConverter converter;
        private AbstractOdfAddin addin;
        private ChainResourceManager resourceManager;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="converter">An implementation of AbstractConverter</param>
        public OdfAddinLib(AbstractOdfAddin addin, AbstractConverter converter)
        {
            this.converter = converter;
            this.addin = addin;
            this.resourceManager = new ChainResourceManager();
            // Add a default resource managers (for common labels)
            this.resourceManager.Add(new System.Resources.ResourceManager("OdfAddinLib.resources.Labels",
                Assembly.GetExecutingAssembly()));
        }

        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager
        {
            set { this.resourceManager.Add(value); }
        }


        /// <summary>
        /// Retrieve the label associated to the specified key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public string GetString(string key)
        {
            return this.resourceManager.GetString(key);
        }

        /// <summary>
        /// Transforms an ODF document into an OOX document.
        /// </summary>
        /// <param name="inputFile">The path of the input file to convert.</param>
        /// <param name="outputFile">The path of the resulting file.</param>
        /// <param name="showUserInterface">True if a progress bar has to be shown, along with a feedback.</param>
        public void OdfToOox(string inputFile, string outputFile, bool showUserInterface)
        {
            if (showUserInterface)
            {
                try
                {
                    // create a temporary file

                    // call the converter
                    using (ConverterForm form = new ConverterForm(this.converter, inputFile, outputFile, this.resourceManager, true))
                    {
                        if (System.Windows.Forms.DialogResult.OK == form.ShowDialog())
                        {
                            if (form.HasLostElements)
                            {
                                string fidelityValue = Microsoft.Win32.Registry.GetValue(this.addin.RegistryKeyUser, ConfigForm.FidelityValue, "false") as string;

                                if (fidelityValue == null)
                                {
                                    fidelityValue = "false";
                                }

                                if (fidelityValue.Equals("false", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, true, "FeedbackLabel", form.LostElements);
                                    infoBox.ShowDialog();
                                }
                            }
                        }
                        else
                        {
                            if (File.Exists(outputFile))
                            {
                                File.Delete(outputFile);
                            }
                            if (form.Exception != null)
                            {
                                throw form.Exception;
                            }
                        }
                    }
                }
                catch (EncryptedDocumentException)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "EncryptedDocumentLabel", "EncryptedDocumentDetail");
                    infoBox.ShowDialog();
                }
                catch (NotAnOdfDocumentException)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "NotAnOdfDocumentLabel", "NotAnOdfDocumentDetail");
                    infoBox.ShowDialog();
                }
                catch (OdfZipUtils.ZipCreationException)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "UnableToCreateOutputLabel", "UnableToCreateOutputDetail");
                    infoBox.ShowDialog();
                }
                catch (OdfZipUtils.ZipException)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "CorruptedInputFileLabel", "CorruptedInputFileDetail");
                    infoBox.ShowDialog();
                }
                catch (Exception e)
                {
                    if (e.InnerException != null && e.InnerException is System.Xml.XmlException)
                    {
                        // An xsl exception may embed an xml exception. In this case we have a non well formed xml document.
                        List<string> messages = new List<string>();
                        messages.Add("CorruptedInputFileDetail");
                        messages.Add("");
                        messages.Add(e.Message);
                        messages.Add("InnerException: " + e.InnerException.Message);
                        InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "CorruptedInputFileLabel", (string[])messages.ToArray());
                        infoBox.ShowDialog();
                    }
                    else
                    {
                        InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "OdfUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");
                        infoBox.ShowDialog();
                    }
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                }
            }
            else
            {
                try
                {
                    converter.DirectTransform = true;
                    converter.Transform(inputFile, outputFile);
                }
                catch (Exception e)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    throw e;
                }
            }
        }

        /// <summary>
        /// Transforms an OOX document into an ODF document.
        /// </summary>
        /// <param name="inputFile">The path of the input file to convert.</param>
        /// <param name="outputFile">The path of the resulting file.</param>
        /// <param name="showUserInterface">True if a progress bar has to be shown, along with a feedback.</param>
        public void OoxToOdf(string inputFile, string outputFile, bool showUserInterface)
        {
            if (showUserInterface)
            {
                try
                {
                    using (ConverterForm form = new ConverterForm(this.converter, inputFile, outputFile, this.resourceManager, false))
                    {
                        System.Windows.Forms.DialogResult dr = form.ShowDialog();

                        if (form.HasLostElements)
                        {
                            string fidelityMsgValue = Microsoft.Win32.Registry.GetValue(this.addin.RegistryKeyUser, ConfigForm.FidelityValue, "false") as string;

                            if (fidelityMsgValue == null || fidelityMsgValue == "false")
                            {
                                InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, true, "FeedbackLabel", form.LostElements);
                                infoBox.ShowDialog();
                            }
                        }
                        if (form.Exception != null)
                        {
                            throw form.Exception;
                        }
                    }
                }
                catch (OdfZipUtils.ZipCreationException zipEx)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "UnableToCreateOutputLabel", zipEx.Message ?? "UnableToCreateOutputDetail");
                    infoBox.ShowDialog();
                }
                catch (OdfZipUtils.ZipException)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "UnableToCreateOutputLabel", "PossiblyEncryptedDocument");
                    infoBox.ShowDialog();
                }
                catch (System.IO.IOException)
                {
                    // this is meant to catch "file already accessed by another process", though there's no .NET fine-grain exception for this.
                    // bug #1676586  Concurrent file access crashes the addin
                    // bug #1807447  avoid display of unlocalized .NET exception message text and display localized string
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "UnableToCreateOutputLabel", "UnableToCreateOutputDetail");
                    //InfoBox infoBox = new InfoBox("UnableToCreateOutputLabel", e.Message, this.resourceManager);
                    infoBox.ShowDialog();
                }
                catch (Exception e)
                {
                    InfoBox infoBox = new InfoBox(this.addin, this.resourceManager, false, "OdfUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");
                    infoBox.ShowDialog();

                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                }
            }
            else
            {
                try
                {
                    converter.DirectTransform = false;
                    converter.Transform(inputFile, outputFile);
                }
                catch (Exception e)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    throw e;
                }
            }
        }

        
        /// <summary>
        /// Build a temporary file.
        /// </summary>
        /// <param name="input">The orginal odf file name</param>
        /// <param name="extension">temp file name extension</param>
        /// <returns>A temporary file name pointing to the user's \Temp folder</returns>
        public string GetTempFileName(string input, string extension)
        {
            // Get the \Temp path
            string tempPath = Path.GetTempPath().ToString();

            // Build the output file name
            string root = null;

            int lastSlash = input.LastIndexOf('\\');
            if (lastSlash > 0)
            {
                root = input.Substring(lastSlash + 1);
            }
            else
            {
                root = input;
            }

            int index = root.LastIndexOf('.');
            if (index > 0)
            {
                root = root.Substring(0, index);
            }

            string output = tempPath + root + "_tmp" + extension;
            int i = 1;

            while (File.Exists(output) || Directory.Exists(output))
            {
                output = tempPath + root + "_tmp" + i + extension;
                i++;
            }
            return output;
        }

        /// <summary>
        /// Create a random temporary folder 
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        /// <param name="targetExtension">The target extension</param>
        /// <returns></returns>
        public string GetTempPath(string fileName, string targetExtension)
        {
            string folderName = null;
            string path = null;
            do
            {
                folderName = Path.GetRandomFileName();
                path = Path.Combine(Path.GetTempPath(), folderName);
            }
            while (Directory.Exists(path));

            Directory.CreateDirectory(path);
            return Path.Combine(path, Path.GetFileNameWithoutExtension(fileName) + targetExtension);
        }

        public void DeleteTempPath(string tempPath)
        {
            string directory = Path.GetDirectoryName(tempPath);
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
    }
}
