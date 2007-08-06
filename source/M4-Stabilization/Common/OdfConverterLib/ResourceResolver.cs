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
using System.Collections;
using System.Text;
using System.Reflection;
using System.Xml;

namespace CleverAge.OdfConverter.OdfConverterLib
{
	/// <summary>
	/// An XmlUrlResolver for embedded resources.
	/// </summary>
	public class ResourceResolver : XmlUrlResolver
    {
        public const string ASSEMBLY_URI_SCHEME = "assembly";
        public const string ASSEMBLY_URI_HOST = "localhost";

		private Assembly assembly;
		private string prefix;
		
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="a">The Assembly instance</param>
		/// <param name="prefix">The prefix used to locate resources within the assembly</param>
		public ResourceResolver(Assembly assembly, string prefix) 
		{
			this.assembly = assembly;
            this.prefix = prefix;
		}
		
		public override Uri ResolveUri(Uri baseUri, String relativeUri)
		{
			if (baseUri == null && !relativeUri.Contains("://"))
            {
                return new Uri(ASSEMBLY_URI_SCHEME + "://" + ASSEMBLY_URI_HOST + "/" + relativeUri);
            }
            else
            {
                return base.ResolveUri(baseUri, relativeUri);
            }
		}
		
		public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
		{
            if (ASSEMBLY_URI_SCHEME.Equals(absoluteUri.Scheme))
            {
                string resource = absoluteUri.OriginalString.Remove(0, ASSEMBLY_URI_SCHEME.Length + ASSEMBLY_URI_HOST.Length + 4).Replace("/", ".");
			    Stream stream = this.assembly.GetManifestResourceStream(this.prefix + "." + resource);
                if (stream != null)
                {
                    return stream;
                }
            }
           
			// use default behaviour instead
			return base.GetEntity(absoluteUri, role, ofObjectToReturn);
		}
	
		public Stream GetInnerStream(String fileName) {
			return this.assembly.GetManifestResourceStream(this.prefix + "." + fileName);
		}
		
	}

}
