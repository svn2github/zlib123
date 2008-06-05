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
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    /// <summary>
    ///     Module ID:      
    ///     Description:    OdfConverter Customization
    ///     Author:         Sateesh
    ///     Create Date:    2008-05-13
    /// </summary>
    public class ConfigManager
    {

        private bool isOox2OdfPackage;

        private bool isErrorIgnored;

        private string configfile;

        public bool IsOox2OdfPackage
        {
            set { isOox2OdfPackage = value; }
            get { return isOox2OdfPackage; }
        }

        public bool IsErrorIgnored
        {
            set { isErrorIgnored = value; }
            get { return isErrorIgnored; }
        }

        public string ConfigFile
        {
            set { configfile = value; }
            get { return configfile; }
        }

        public ConfigManager(string filename)
        {
            configfile = filename;
            isOox2OdfPackage = true;
            isErrorIgnored = false;
        }

        public void LoadConfig()
        {
            XmlTextReader reader = null;
            bool isConfigOneFound = false;
            bool isConfigTwoFound = false;
            string exceptionString = "";
            try
            {
                reader = new XmlTextReader(configfile);
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            {
                                if (reader.Name.Equals("oox2odf_package"))
                                {
                                    this.isOox2OdfPackage = Convert.ToBoolean(reader.ReadString());
                                    isConfigOneFound = true;
                                }
                                else if (reader.Name.Equals("ignore_error"))
                                {
                                    this.isErrorIgnored = Convert.ToBoolean(reader.ReadString());
                                    isConfigTwoFound = true;
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                reader.Close();
            }
           
        }

        public void SaveConfig()
        {
            XmlTextWriter writer = null;
            try
            {
                writer = new XmlTextWriter(configfile, Encoding.UTF8);
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteComment(" OdfConverter ");
                writer.WriteStartElement("Configuration");
                writer.WriteElementString("oox2odf_package", Convert.ToString(isOox2OdfPackage));
                writer.WriteElementString("ignore_error", Convert.ToString(isErrorIgnored));
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                writer.Close();
            }
        }
    }
}
