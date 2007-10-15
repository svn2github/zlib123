using System;
using System.Runtime.InteropServices;
using System.Text;

namespace OdfConverterHost {
    abstract class Win32 {

        [DllImport("kernel32.dll")]
        public static extern void OutputDebugString(string s);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", EntryPoint="PostMessageW")]
        public static extern int PostMessage(IntPtr hWnd, int nMsg, int wParam, int lParam);

        public const int HCBT_CREATEWND = 3;
        public const int HCBT_ACTIVATE = 5;

        public const int WH_CBT = 5;


        public const int SWP_NOSIZE = 1;
        public const int SWP_NOMOVE = 2;
        public const int SWP_NOZORDER = 4;

        public const int GWL_STYLE = -16;
        public const int WS_CHILD = 0x40000000;

        public const int WS_EX_TOPMOST = 0x00000008;

        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public class CREATESTRUCT {
            private IntPtr _ptr;
            public CREATESTRUCT(IntPtr ptr) {
                _ptr = ptr;
            }

            public IntPtr lpCreateParams {
                get {
                    return Marshal.ReadIntPtr(_ptr);
                }
            }

            public IntPtr hInstance {
                get {
                    return Marshal.ReadIntPtr(_ptr, 4);
                }
            }

            public IntPtr hMenu {
                get {
                    return Marshal.ReadIntPtr(_ptr, 8);
                }
            }

            public IntPtr hwndParent {
                get {
                    return Marshal.ReadIntPtr(_ptr, 12);
                }
            }

            public int cy {
                get {
                    return Marshal.ReadInt32(_ptr, 16);
                }
                set {
                    Marshal.WriteInt32(_ptr, 16, value);
                }
            }
            public int cx {
                get {
                    return Marshal.ReadInt32(_ptr, 20);
                }
                set {
                    Marshal.WriteInt32(_ptr, 20, value);
                }
            }
            public int y {
                get {
                    return Marshal.ReadInt32(_ptr, 24);
                }
                set {
                    Marshal.WriteInt32(_ptr, 24, value);
                }
            }
            public int x {
                get {
                    return Marshal.ReadInt32(_ptr, 28);
                }
                set {
                    Marshal.WriteInt32(_ptr, 28, value);
                }
            }
            public uint style {
                get {
                    return (uint)Marshal.ReadInt32(_ptr, 32);
                }
            }
            public IntPtr lpszName {
                get {
                    return Marshal.ReadIntPtr(_ptr, 36);
                }
            }
            public string Name {
                get {
                    IntPtr ptrName = lpszName;
                    if (ptrName != IntPtr.Zero) {
                        return Marshal.PtrToStringAnsi(ptrName);
                    } else {
                        return "(no name)";
                    }
                }
            }
            public IntPtr lpszClass {
                get {
                    return Marshal.ReadIntPtr(_ptr, 40);
                }
            }
            public int dwExStyle {
                get {
                    return Marshal.ReadInt32(_ptr, 44);
                }
                set {
                    Marshal.WriteInt32(_ptr, 44, value);
                }
            }

            public bool HasStyle(uint nStyle) {
                uint myStyle = this.style;
                return (myStyle & nStyle) == nStyle;
            }

        }

        public class CBT_CREATEWND {
            private IntPtr _ptr;
            public CBT_CREATEWND(IntPtr ptr) {
                _ptr = ptr;
            }
            public CREATESTRUCT lpcs {
                get {
                    return new CREATESTRUCT(Marshal.ReadIntPtr(_ptr));
                }
            }
                public IntPtr hwndInsertAfter {
                    get {
                        return Marshal.ReadIntPtr(_ptr, 4);
                    }
                }
            }

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc proc, IntPtr hInstance, uint dwThreadId);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int UnhookWindowsHookEx(IntPtr hhook);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(IntPtr hhook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int GetWindowRect(IntPtr hWnd, ref System.Drawing.Rectangle rect);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder buffer, int len);

        public static string GetWindowText(IntPtr hWnd) {
            StringBuilder builder = new StringBuilder(512);
            int nLen = GetWindowText(hWnd, builder, builder.Capacity);
            builder.Length = nLen;
            return builder.ToString();
        }


        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);


        public static int GetWindowStyle(IntPtr hWnd) {
            return GetWindowLong(hWnd, GWL_STYLE);
        }

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hWnd);
    }
}
