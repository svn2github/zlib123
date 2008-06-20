/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ``AS IS´´ AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;

namespace OdfConverter.OdfConverterLib.ManagedHelpers
{
    // This interface will be implemented by the outer object in the
    // aggregation - that is, by the shim.
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("7B70C487-B741-4973-B915-B812A91BDF63")]
    internal interface IComAggregator
    {
        void SetInnerPointer(IntPtr pUnkInner);
    }

    // This interface is implemented by the managed aggregator - the single
    // method is a wrapper around Marshal.CreateAggregatedObject, which can be
    // called from unmanaged code (that is, called from the shim).
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("142A261B-1550-4849-B109-715AA4629A14")]
    internal interface IManagedAggregator
    {
        void CreateAggregatedInstance(
            string assemblyName, string typeName, IComAggregator outerObject);
    }

    // The unmanaged shim will instantiate this object in order to call
    // through to Marshal.CreateAggregatedObject.
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ManagedHelpers.ManagedAggregator")]
    internal class ManagedAggregator : IManagedAggregator
    {
        public void CreateAggregatedInstance(
            string assemblyName, string typeName, IComAggregator outerObject)
        {
            IntPtr pOuter = IntPtr.Zero;
            IntPtr pInner = IntPtr.Zero;

            try
            {
                // We use Marshal.CreateAggregatedObject to create a CCW where
                // the inner object (the target managed add-in) is aggregated 
                // with the supplied outer object (the shim).
                pOuter = Marshal.GetIUnknownForObject(outerObject);
                object innerObject =
                    AppDomain.CurrentDomain.CreateInstanceAndUnwrap(
                    assemblyName, typeName);
                pInner = Marshal.CreateAggregatedObject(pOuter, innerObject);

                // Make sure the shim has a pointer to the add-in.
                outerObject.SetInnerPointer(pInner);
            }
            finally
            {
                if (pOuter != IntPtr.Zero)
                {
                    Marshal.Release(pOuter);
                }
                if (pInner != IntPtr.Zero)
                {
                    Marshal.Release(pInner);
                }
            }
        }
    }
}
