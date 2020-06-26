using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System.ServiceProcess;
using Microsoft.Win32;
using Gma.System.MouseKeyHook;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace PunchClock
{
    class Program
    {
       
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static DateTime strikeStartKeyboard;
        private static DateTime strikeStartMouse;
        private static long strikeCount = 0;
        static void Main(string[] args)
        {
            string MachineName1 = Environment.MachineName;
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            // GetLastLoginToMachine(MachineName1, userName);

            //ExportToExcel();
            //DateTime WorkStart = DateTime.Now;
            ////startupMaker();
            //var handle = GetConsoleWindow();
            //ListenForMouseEvents();
            //// Hide
            //ShowWindow(handle, SW_HIDE);
            //_hookID = SetHook(_proc);
            Application.Run();
            //UnhookWindowsHookEx(_hookID);
            UpTime();
        }
        private static void startupMaker()
        {
            RegistryKey re = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            var ss = System.Windows.Forms.Application.ExecutablePath.ToString();
            re.SetValue("ConsoleApp2", System.Windows.Forms.Application.ExecutablePath.ToString());

        }
        public static void ListenForMouseEvents()
        {
            Console.WriteLine("Listening to mouse clicks.");

            //When a mouse button is pressed 
            Hook.GlobalEvents().MouseDown += async (sender, e) =>
            {
                try
                {
                    if ((DateTime.Now - strikeStartKeyboard).Minutes > 15)
                    {
                        for (var i = 0; (DateTime.Now - strikeStartMouse).Minutes / 15 > i; i++)
                        {
                            strikeCount++;
                        }
                    }
                }
                catch (Exception wss) { }
            };
            //When a double click is made
            Hook.GlobalEvents().MouseDoubleClick += async (sender, e) =>
            {
                try
                {
                    if ((DateTime.Now - strikeStartKeyboard).Minutes > 15)
                    {
                        for (var i = 0; (DateTime.Now - strikeStartMouse).Minutes / 15 > i; i++)
                        {
                            strikeCount++;
                        }
                    }
                }
                catch (Exception ss) { }
            };
        }
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
            }
        }
        private delegate IntPtr LowLevelKeyboardProc(
        int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(
            int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                //int vkCode = Marshal.ReadInt32(lParam);
                //Console.WriteLine((Keys)vkCode);
                //StreamWriter sw = new StreamWriter(Application.StartupPath + @"\log.txt", true);
                //sw.Write((Keys)vkCode+"-");
                //sw.Close();
                var ticks = Stopwatch.GetTimestamp();
                var uptime = ((double)ticks) / Stopwatch.Frequency;
                var uptimeSpan = TimeSpan.FromSeconds(uptime);
                if ((DateTime.Now - strikeStartMouse).Minutes>15) {
                    for (var i = 0; (DateTime.Now - strikeStartKeyboard).Minutes / 15 > i; i++)
                    {
                        strikeCount++;
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        //These Dll's will handle the hooks. Yaaar mateys!

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        // The two dll imports below will handle the window hiding.

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        public static DateTime? GetLastLoginToMachine(string machineName, string userName)
        {
            EventLog myLog = new EventLog();
            myLog.Log = "System";
            myLog.Source = "User32";

            var lastEntry = myLog;
            EventLogEntry sw;
            for (var i = myLog.Entries.Count -1 ; i >1; i--)
            {
                if (lastEntry.Entries[i].InstanceId == 1074)
                {
                    sw = lastEntry.Entries[i];
                    break;
                }
            }
           // var last_error_Message = lastEntry.Message;

            //for (int index = myLog.Entries.Count - 1; index > 0; index--)
            //{
            //    var errLastEntry = myLog.Entries[index];
            //    if (errLastEntry.EntryType == EventLogEntryType.Error)
            //    {
            //        //this is the last entry with Error
            //        var appName = errLastEntry.Source;
            //        break;
            //    }
            //}
            return null;

        }
        public static DateTime GetLastSystemShutdown()
        {
            string sKey = @"System\CurrentControlSet\Control\Windows";
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sKey);

            string sValueName = "ShutdownTime";
            byte[] val = (byte[])key.GetValue(sValueName);
            long valueAsLong = BitConverter.ToInt64(val, 0);
            var ss =  DateTime.FromFileTime(valueAsLong);
            return ss;
        }
        public static TimeSpan UpTime()
        {
                using (var uptime = new PerformanceCounter("System", "System Up Time"))
                {
                    uptime.NextValue();       //Call this an extra time before reading its value
                    return TimeSpan.FromSeconds(uptime.NextValue());
                }
            
        }
        protected static void ExportToExcel()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "4";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "5";
            xlWorkSheet.Cells[3, 2] = "Two";



            xlWorkBook.SaveAs(Application.StartupPath + @"\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
