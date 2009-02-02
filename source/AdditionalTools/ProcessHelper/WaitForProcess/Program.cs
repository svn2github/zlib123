/* 
 * Copyright (c) 2009 DIaLOGIKa
 *
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
using System.Diagnostics;
using System.Threading;
using System.Management;

namespace WaitForProcess
{
    /// <summary>
    /// A simply utility which makes the calling process wait until a specified process has exited.
    /// </summary>
    /// <remarks>
    /// With the .NET 3.5 SP1 update Microsoft has fixed a problem in the Windows SDK regarding UAC in Vista, 
    /// however they made things unfortunately worse for self-extracting .exe scenarios as in our case. 
    /// 
    /// Microsoft provides no workaround to this problem and a fix won’t be available before VS 2010. 
    /// That means we have to consider another solution for our setup.
    ///
    /// The solution implemented now is that we start this utility in IExpress' post-install command to make
    /// IExpress wait for the installation to complete.
    /// 
    /// Details
    /// =======
    /// The setup procedure works as follows:
    /// 
    /// 1.	OdfAddInForOfficeSetup-en.exe is a self-extracting installer created with IExpress
    /// 2.	IExpress extracts the file to a temporary folder (%TMP%\IXPnnn.TMP)
    /// 3.	IExpress launches the actual setup command which has been specified, i.e. start /W setup.exe
    /// 4.	On return of setup.exe, IExpress cleans up the temporary folder created in step 2.
    /// 
    /// This used to work fine because when the setup.exe process returned, it was safe to delete
    /// all temporary files. However, the change introduced in VS 2008 SP1 causes the setup.exe process 
    /// to return immediately, leading to a race condition where the temporary files 
    /// might get deleted before the installation is completed.
    /// 
    /// References
    /// ==========
    /// https://connect.microsoft.com/VisualStudio/feedback/ViewFeedback.aspx?FeedbackID=369138 
    /// http://social.msdn.microsoft.com/Forums/en-US/winformssetup/thread/3731985c-d9cc-4403-ab7d-992a0971f686/?ffpr=0 
    /// </remarks>
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: WaitForProcess.exe <process name> <window title>");
            }
            else
            {
                // We wait to make sure that the msiexec process has already been launched
                Thread.Sleep(5000);

                string processName = args[0];
                string installerName = args[1];

                try
                {
                    // use WMI to retrieve the command line
                    SelectQuery selectQuery = new SelectQuery(string.Format("select CommandLine, ProcessId from Win32_Process where name='{0}'", processName));

                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(selectQuery))
                    {
                        foreach (ManagementObject wmiProcess in searcher.Get())
                        {
                            try
                            {
                                string commandLine = wmiProcess.Properties["CommandLine"].Value.ToString();

                                // check whether we got the right process where our installer is contained in the command-line args
                                if (commandLine.ToLowerInvariant().Contains(installerName.ToLowerInvariant()))
                                {
                                    // we assume the process id to be numeric. if it isn't we are pretty much out of luck anyway
                                    int processId = int.Parse(wmiProcess.Properties["ProcessId"].Value.ToString());
                                    
                                    Process process = Process.GetProcessById(processId);
                                    if (process != null && !process.HasExited)
                                    {
                                        // wait for the installer to complete
                                        process.WaitForExit();
                                        return;
                                    }
                                }

                            }
                            catch (Exception)
                            {
                                // fail silently (yes, we do, hehe)
                            }

                        }
                    }
                }
                catch (Exception)
                {
                }
                
                // code using System.Diagnostics only
                //Process[] processList = Process.GetProcessesByName(processName);

                //foreach (Process process in processList)
                //{
                //    if (process.MainWindowTitle.ToLowerInvariant().Contains(installerName.ToLowerInvariant()))
                //    {
                //        // max wait 1h
                //        process.WaitForExit(3600000);
                //        Console.WriteLine("Installation finished.");
                //        return;
                //    }
                //}
            }
        }
    }
}
