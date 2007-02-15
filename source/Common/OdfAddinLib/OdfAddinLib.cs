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

namespace CleverAge.OdfConverter.OdfConverterLib
{

    /// <summary>
    /// Chained resource managers
    /// </summary>
    internal class AddinResourceManager : System.Resources.ResourceManager
    {
        private System.Resources.ResourceManager primaryResourceManager;

        /// <summary>
        /// Default resource manager
        /// </summary>
        /// <param name="path"></param>
        /// <param name="assembly"></param>
        public AddinResourceManager(string path, Assembly assembly) : base (path, assembly)
        { }

        public System.Resources.ResourceManager PrimaryResourceManager
        {
            set { this.primaryResourceManager = value; }
        }

        public override string GetString(string name)
        {
            if (this.primaryResourceManager != null)
            {
                string s = this.primaryResourceManager.GetString(name);   
                if (s != null && s.Length > 0)
                {
                    return s;
                }
            }
            return base.GetString(name);
        }


    }

    public class OdfAddinLib
    {
        private AbstractConverter converter;
        private System.Resources.ResourceManager resourceManager;

        public OdfAddinLib(AbstractConverter converter)
        {
            this.converter = converter;
            this.resourceManager = 
                new AddinResourceManager("OdfAddinLib.resources.Labels", 
                Assembly.GetExecutingAssembly());
        }

        public System.Resources.ResourceManager OverrideResourceManager
        {
            set { ((AddinResourceManager) this.resourceManager).PrimaryResourceManager = value; }
        }


        /// <summary>
        /// Returns the ResourceManager containing the labels of the application.
        /// </summary>
        /// <returns>The ResourceManager.</returns>
        public string GetString(string name)
        {
            return this.resourceManager.GetString(name);
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
                                ArrayList elements = form.LostElements;
                                InfoBox infoBox = new InfoBox("FeedbackLabel", elements, this.resourceManager);
                                infoBox.ShowDialog();
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
                    InfoBox infoBox = new InfoBox("EncryptedDocumentLabel", "EncryptedDocumentDetail", this.resourceManager);
                    infoBox.ShowDialog();
                }
                catch (NotAnOdfDocumentException)
                {
                    InfoBox infoBox = new InfoBox("NotAnOdfDocumentLabel", "NotAnOdfDocumentDetail", this.resourceManager);
                    infoBox.ShowDialog();
                }
                catch (OdfZipUtils.ZipCreationException)
                {
                    InfoBox infoBox = new InfoBox("UnableToCreateOutputLabel", "UnableToCreateOutputDetail", this.resourceManager);
                    infoBox.ShowDialog();
                }
                catch (Exception e)
                {
                    InfoBox infoBox = new InfoBox("OdfUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")", this.resourceManager);
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
                            ArrayList elements = form.LostElements;
                            InfoBox infoBox = new InfoBox("FeedbackLabel", elements, this.resourceManager);
                            infoBox.ShowDialog();
                        }

                        if (form.Exception != null)
                        {
                            throw form.Exception;
                        }
                    }
                } catch (OdfZipUtils.ZipCreationException zipEx) {
                    InfoBox infoBox = new InfoBox("UnableToCreateOutputLabel", zipEx.Message ?? "UnableToCreateOutputDetail", this.resourceManager);
                    infoBox.ShowDialog();
                } 
                catch (Exception e)
                {
                    InfoBox infoBox = new InfoBox("OdfUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")", this.resourceManager);
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
        /// Returns the logo of the application.
        /// </summary>
        /// <returns>The logo of the application.</returns>
        public stdole.IPictureDisp GetLogo()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("OdfLogo.png"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }
            }
            if (stream == null)
            {
                return null;
            }
            System.Drawing.Bitmap image = new System.Drawing.Bitmap(stream);
            return OdfAddinLib.ConvertImage.Convert(image);
        }

        /// <summary>
        /// Build a temporary file.
        /// </summary>
        /// <param name="input">The orginal odf file name</param>
        /// <returns>A temporary file name pointing to the user's \Temp folder</returns>
        public string GetTempFileName(string input)
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

            string output = tempPath + root + "_tmp.docx";
            int i = 1;

            while (File.Exists(output) || Directory.Exists(output))
            {
                output = tempPath + root + "_tmp" + i + ".docx";
                i++;
            }
            return output;
        }


        sealed private class ConvertImage : System.Windows.Forms.AxHost
        {
            private ConvertImage()
                : base(null)
            {
            }
            public static stdole.IPictureDisp Convert
                (System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)System.
                    Windows.Forms.AxHost
                    .GetIPictureDispFromPicture(image);
            }
        }
    }
}
