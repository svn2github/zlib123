using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CleverAge.OdfConverter.OdfWord2003Addin
{
    public partial class FrmPrerequisites : Form
    {
        [StructLayout(LayoutKind.Sequential)]
        private  struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public int Width {
                get {
                    return Right - Left;
                }
            }
            public int Height {
                get {
                    return Bottom - Top;
                }
            }
        }


        [DllImport("user32.dll")]
        private extern static IntPtr FindWindow(IntPtr lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private extern static bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);
        public FrmPrerequisites() {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        private extern static int ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_SHOW = 1;
        private const int SW_HIDE = 0;

        IntPtr _hParent;

        private void Form1_Load(object sender, EventArgs e) {
            _hParent = FindWindow(IntPtr.Zero, "ODF Add-in for Word 2003");
            if (_hParent != IntPtr.Zero) {
                RECT rc = new RECT();
                if (GetWindowRect(_hParent, ref rc)) {
                    this.Location = new Point(rc.Left, rc.Top);
                    this.Size = new Size(rc.Width, rc.Height);
                }
                ShowWindow(_hParent, SW_HIDE);
            }
            // Must be set AFTER FindWindow otherwise FindWindow will find ourselve...
            this.Text = "ODF Add-in for Word 2003";
        }

        private void linkAddRemove_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start("RunDll32.exe","shell32.dll,Control_RunDLL appwiz.cpl");

        }

        private void linkOfficePIA_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start("http://www.microsoft.com/downloads/details.aspx?familyid=3c9a983a-ac14-4125-8ba0-d36d67e0f4ad&displaylang=en");
        }

        private void FrmPrerequisites_FormClosing(object sender, FormClosingEventArgs e) {
            ShowWindow(_hParent, SW_SHOW);
        }



    }
}