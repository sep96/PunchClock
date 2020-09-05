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
using System.Data.SqlClient;
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
        private static DateTime WorkStart;
        private static long strikeCount = 0;
        private static readonly string connectionString = "Data Source=.;Initial Catalog=PC;Integrated Security=True";
        private static string userName = "";
        static void Main(string[] args)
        {
            string MachineName1 = Environment.MachineName;
            userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            GetLastLoginToMachine(MachineName1, userName);
            WorkStart = DateTime.Now;
            var firstDayOfMonth = new DateTime(WorkStart.Year, WorkStart.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            if (lastDayOfMonth.Day == WorkStart.Day)
            {
                ExportMonth();
            }
            if (5 == WorkStart.Day)
            {
                ExportMonth();
            }
            strikeStartKeyboard = DateTime.Now;
            ////startupMaker();
            var handle = GetConsoleWindow();
            ListenForMouseEvents();
            //// Hide
            ShowWindow(handle, SW_HIDE);
            _hookID = SetHook(_proc);
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
                var ticks = Stopwatch.GetTimestamp();
                var uptime = ((double)ticks) / Stopwatch.Frequency;
                var uptimeSpan = TimeSpan.FromSeconds(uptime);
                var diffrenceTime = (DateTime.Now - strikeStartKeyboard).Minutes / 1;
                //if ((DateTime.Now - strikeStartMouse).Minutes>15) {
                    for (var i = 0; diffrenceTime > i; i++)
                    {
                        string query = " insert into  [PC].[dbo].[TimeRecorder] ([username],[Gap] ,[Date],[Start]) VALUES (@user , @time , @date , @start)";
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            using(SqlCommand comm = new SqlCommand(query, conn))
                            {
                                comm.Parameters.AddWithValue("@user", userName);
                                comm.Parameters.AddWithValue("@time", "15");
                                comm.Parameters.AddWithValue("@date", DateTime.Now);
                                comm.Parameters.AddWithValue("@start", WorkStart);
                                comm.ExecuteNonQuery();
                            }
                            conn.Close();
                        }
                    }
              //  }
            }

            strikeStartKeyboard = DateTime.Now;
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
            string sKey = @"System\CurrentControlSet\Control\Windows";
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sKey);
            string sValueName = "ShutdownTime";
            byte[] val = (byte[])key.GetValue(sValueName);
            long valueAsLong = BitConverter.ToInt64(val, 0);
            var result = DateTime.FromFileTime(valueAsLong);
            string query = " update  [PC].[dbo].[TimeRecorder] set [EndTime]=CONVERT(datetime,@end)  where CONVERT(DATE,[Start])=CONVERT(DATE,@end)";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand comm = new SqlCommand(query, conn))
                {
                    comm.Parameters.AddWithValue("@end", result);
                    comm.ExecuteNonQuery();
                }
                conn.Close();
            }
            return result;
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
        private static void ExportMonth()
        {
            string query = "SELECT[username]  , sum ([Gap]) " +
                           "FROM [PC].[dbo].[TimeRecorder] where MONTH([Date]) = @month and YEAR([Date]) = @year and [username]=@user " +
                           " group by [username];";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand comm = new SqlCommand(query, conn))
                {
                    comm.Parameters.AddWithValue("@month", DateTime.Now.Month);
                    comm.Parameters.AddWithValue("@year", DateTime.Now.Year);
                    comm.Parameters.AddWithValue("@user", userName);
                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            StreamWriter sw = new StreamWriter(Application.StartupPath + @"\"+userName+"-" +DateTime.Now.Year + DateTime.Now.Month + ".txt", true);
                            sw.Write(reader.GetString(0) + "-" + reader.GetInt32(1));
                            sw.Close();
                            strikeCount++;
                        }
                    }
                }
                conn.Close();
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
