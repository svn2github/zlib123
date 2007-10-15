using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;

namespace OdfConverterHost {
    static class Program {
        public static Form1 MainForm;

        private static int _converterCookie = 0;



        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            RegistrationServices rs = new RegistrationServices();
            try {
                Win32.OutputDebugString("OdfConverterHost launched, ThreadId = " + Win32.GetCurrentThreadId().ToString("x") + "\n");

                // Register "Single use" : one client only => the process will shutdown when the client leaves
                _converterCookie = rs.RegisterTypeForComClients(typeof(Converter),
                                                              RegistrationClassContext.LocalServer,
                                                              RegistrationConnectionType.SingleUse);
                Win32.OutputDebugString("Class registered\n");


#if false
                MainForm = new Form1();
                Application.Run(MainForm);
#else
                Application.Run();
#endif
                Win32.OutputDebugString("Run terminated\n");
            } catch (Exception ex) {
                MessageBox.Show(ex.StackTrace, ex.Message);
            } finally {
                // Cleanup... without throwing anything
                try {
                    if (_converterCookie != 0) {
                        rs.UnregisterTypeForComClients(_converterCookie);
                    }
                } catch {
                    // Do nothing
                }
            }
            Win32.OutputDebugString("OdfConverterHost exiting\n");
        }

    }
}