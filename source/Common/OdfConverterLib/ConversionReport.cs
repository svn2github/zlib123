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
using System.IO;

namespace CleverAge.OdfConverter.OdfConverterLib
{
    public class ConversionReport
    {
        public const int DEBUG_LEVEL = 1;
        public const int INFO_LEVEL = 2;
        public const int WARNING_LEVEL = 3;
        public const int ERROR_LEVEL = 4;

        private StreamWriter writer = null;
        private int level = INFO_LEVEL;

        public ConversionReport(string filename, int level)
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
