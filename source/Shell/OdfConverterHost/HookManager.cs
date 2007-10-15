using System;
using System.Collections.Generic;
using System.Text;

namespace OdfConverterHost {
    class HookManager : IDisposable {

        #region Private Fields
        private Win32.HookProc _hookProc;
        private IntPtr _hhook = IntPtr.Zero;
        private int _x;
        private int _y;
        #endregion

        #region Construction
        public HookManager(int centerPos) {
            // Install hook
            _hookProc = new Win32.HookProc(MyCbtProc);
            _hhook = Win32.SetWindowsHookEx(Win32.WH_CBT, _hookProc, IntPtr.Zero, Win32.GetCurrentThreadId());
            _x = centerPos / 65536;
            _y = centerPos % 65536;
        }
        #endregion

        #region Hook handling
        private int MyCbtProc(int nCode, IntPtr wParam, IntPtr lParam) {
            if (nCode < 0) {
                return Win32.CallNextHookEx(_hhook, nCode, wParam, lParam);
            } else {
                switch (nCode) {
                case Win32.HCBT_CREATEWND: {
                        Win32.OutputDebugString("HCBT_CREATEWND : " + Win32.GetWindowText(wParam) + "\n");
                        Win32.CBT_CREATEWND cbtCreateWindow = new Win32.CBT_CREATEWND(lParam);
                        Win32.CREATESTRUCT lpcs = cbtCreateWindow.lpcs;
                        if (!lpcs.HasStyle(Win32.WS_CHILD)) {
                            Win32.OutputDebugString("   name = " + lpcs.Name + "\n");

                            lpcs.x = _x - (lpcs.cx / 2);
                            lpcs.y = _y - (lpcs.cy / 2);

                            lpcs.dwExStyle |= Win32.WS_EX_TOPMOST;
                        }
                    }
                    break;
                default:
                    Win32.OutputDebugString("MyCbtProc : " + nCode.ToString() + "\n");
                    break;
                }
                Win32.CallNextHookEx(_hhook, nCode, wParam, lParam);
                return 0;
            }
        }
        #endregion



        #region IDisposable Members
        private bool _disposed = false;
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~HookManager() {
            Dispose(false);
        }
        protected void Dispose(bool disposing) {
            if (!_disposed) {
                if (disposing) {
                    // Dispose managed ressource => none
                }
                // Dispose unmanaged ressources
                if (_hhook != IntPtr.Zero) {
                    Win32.UnhookWindowsHookEx(_hhook);
                }

                _disposed = true;
            }

        }
        #endregion
    }
}
