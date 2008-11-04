/*
 * Copyright (c) 2008, DIaLOGIKa
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
using System.Reflection;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using OdfConverter.Office;

namespace OdfConverter.OdfConverterLib
{
    [ComVisible(false)]
    public class LateBindingObject : IDisposable
    {
        protected object _obj;

        public LateBindingObject(object obj)
        {
            _obj = obj;
        }

        public LateBindingObject Invoke(string name, params object[] parameters)
        {
            object obj = _obj.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, _obj, parameters);
            return new LateBindingObject(obj);
        }

        public EventInfo GetEvent(string eventName)
        {
            return _obj.GetType().GetEvent(eventName);
        }

        private IConnectionPoint _cp;
        private int _dwCookie = 0;
        
        public void AddClickEventHandler(Office.CommandBarButtonEvents_ClickEventHandler handler) 
        {
            IConnectionPointContainer cpc;
            CommandBarButtonEvents sink;
            try
            {
                cpc = _obj as IConnectionPointContainer;
                if (cpc == null)
                    return;
                Guid iid = typeof(ICommandBarButtonEvents).GUID;
                cpc.FindConnectionPoint(ref iid, out _cp);
                if (_cp == null)
                    return;
                sink = new CommandBarButtonEvents();
                sink.Register(handler as CommandBarButtonEvents_ClickEventHandler);
                _cp.Advise(sink, out _dwCookie);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            finally
            {
                sink = null;
                cpc = null;
            }
        }

#region Methods to access properties
        protected object Get(string propertyName)
        {
            return _obj.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, _obj, null);
        }

        protected void Set(string propertyName, object value)
        {
            _obj.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, _obj, new object[] { value });
        }

        public Int32 GetInt32(string propertyName)
        {
            return Convert.ToInt32(this.Get(propertyName));
        }

        public void SetInt32(string propertyName, Int32 value)
        {
            object obj = value;
            Set(propertyName, obj);
        }

        public bool GetBool(string propertyName)
        {
            return Convert.ToBoolean(this.Get(propertyName));
        }

        public void SetBool(string propertyName, bool value)
        {
            object obj = value;
            Set(propertyName, obj);
        }

        public string GetString(string propertyName)
        {
            return Convert.ToString(this.Get(propertyName));
        }

        public void SetString(string propertyName, string value)
        {
            object obj = value;
            Set(propertyName, obj);
        }
#endregion

        public Int32 ToInt32()
        {
            return Convert.ToInt32(_obj);
        }

        public override string ToString()
        {
            return Convert.ToString(_obj);
        }

        public void Dispose()
        {
            if (_obj != null)
            {
                Marshal.ReleaseComObject(_obj);
                _obj = null;
            }
        }
    }
}
