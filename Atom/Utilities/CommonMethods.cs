using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Atom.Utilities
{
    class CommonMethods
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// Getting the latest Directory from a  Parent Directory
        /// </summary>
        /// <param name="ParentDirectory"></param>
        /// <param name="TestContext"></param>
        /// <returns></returns>
        public string GetLatestFolderLocation(string ParentDirectory)
        {
            string LatestDirectory = "";
            try
            {
                DirectoryInfo parentDirectorInfo = new DirectoryInfo(ParentDirectory);
                if (parentDirectorInfo.Exists)
                {
                    //Find the latet directory
                    DirectoryInfo LatestDir = parentDirectorInfo.GetDirectories()
                           .OrderByDescending(d => d.LastWriteTimeUtc).First();
                    LatestDirectory = ParentDirectory + @"\" + LatestDir.Name.ToString();
                }
                else
                    Trace.WriteLine("Parent Path " + ParentDirectory + " does not exist/ not accessible");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Error Occured in getting the access of directory! ERROR: " + e.Message);
                //TestContext.WriteLine("Error Occured in getting the access of directory! ERROR: " + e.Message);
            }
            return LatestDirectory;
        }

        /// <summary>
        /// Kills all running browser processes
        /// </summary>
        public static void KillBrowserProcesses()
        {
            try
            {
                //Process RequiredProcess = null;
                if (Process.GetProcessesByName("chrome").Any())
                {
                    Process[] processes = Process.GetProcessesByName("chrome");
                    foreach (Process proc in processes)
                    {
                        proc.Kill();
                    }
                }

                if (Process.GetProcessesByName("iexplore").Any())
                {
                    Process[] processes = Process.GetProcessesByName("iexplore");
                    foreach (Process proc in processes)
                    {
                        proc.Kill();
                    }
                }

                if (Process.GetProcessesByName("IEDriverServer").Any())
                {
                    Process[] processes = Process.GetProcessesByName("IEDriverServer");
                    foreach (Process proc in processes)
                    {
                        proc.Kill();
                    }
                }

                if (Process.GetProcessesByName("chromedriver").Any())
                {
                    Process[] processes = Process.GetProcessesByName("chromedriver");
                    foreach (Process proc in processes)
                    {
                        proc.Kill();
                    }
                }
            }
            catch
            {

            }
        }

        /// <summary>
        /// Method for bringing the IE window in front for runing the test and avoiding the error when wondow is not focused
        /// </summary>
        public void BringWindowInFront()
        {
            //Process RequiredProcess = null;
            Process[] list = Process.GetProcesses();
            if (Process.GetProcessesByName("chrome").Any())
            {
                Process[] processes = Process.GetProcessesByName("chrome");
                foreach (Process proc in processes)
                {
                    if (proc.MainWindowTitle.Contains("google"))
                    {
                        //TestContext.WriteLine("FOUND THE IE WINDOW...");
                        IntPtr pointer = proc.Handle;
                        SetForegroundWindow(pointer);
                        //RequiredProcess = proc;
                        //proc.WaitForInputIdle();
                        //SendMessage(pointer, WM_SYSCOMMAND, SC_RESTORE, 0);
                        IntPtr hWnd = proc.Handle;
                        //SetFocus(new HandleRef(null, hWnd));
                    }
                }
            }
        }

        /// <summary>
        /// Logs input text into additional information section of TRX file
        /// </summary>
        /// <param name="textToLog"></param>
        /// <param name="testContext"></param>
        public static void LogTextIntoAdditionalInformationSection(string textToLog, TestContext testContext)
        {
            testContext.WriteLine(textToLog);
        }

        /// <summary>
        /// generate n digit random number
        /// </summary>
        /// <param name="digCount"></param>
        /// <returns></returns>
        public String GetRandomNumber(int digCount)
        {
            Random rnd = new Random();
            StringBuilder sb = new StringBuilder(digCount);
            for (int i = 0; i < digCount; i++)
                sb.Append((char)('0' + rnd.Next(1, 10)));
            return sb.ToString();
        }
    }
}
