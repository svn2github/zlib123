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
