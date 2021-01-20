using IWshRuntimeLibrary;
using OutlookReminderTopMost.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookReminderTopMost
{
    class Program
    {


        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_NOZORDER = 0x0004;
        const UInt32 SWP_SHOWWINDOW = 0x0040;
        


        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        protected delegate bool MrEnumWindowsCallback(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        protected static extern bool EnumWindows(MrEnumWindowsCallback enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        protected static extern bool IsWindowVisible(IntPtr hWnd);

        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        private static Process[] s_processes;
        private static string[] s_searchedTitles;

        static void Main(string[] args)
        {
            try
            {
                string binary = Process.GetCurrentProcess().MainModule.FileName;
                s_searchedTitles = Settings.Default.NamesOfSearchedHeaders?.Split(';').Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToArray();
                if (s_searchedTitles == null || s_searchedTitles.Length == 0)
                {
                    throw new Exception(string.Format("Missing '{0}' in configuration file '{1}'", nameof(Settings.Default.NamesOfSearchedHeaders), binary + ".config"));
                }

                string autostartDir = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
                string lnk = Path.Combine(autostartDir, Path.GetFileNameWithoutExtension(binary) + ".lnk");


                if (!System.IO.File.Exists(lnk))
                {
                    // Create shortcut
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnk);

                    shortcut.Description = "Outlook Reminder Top Most";   // The description of the shortcut
                    shortcut.TargetPath = binary;                 // The path of the file that will launch when the shortcut is run
                    shortcut.Save();
                }

                Stopwatch lastStartTry = null;
                while (true)
                {
                    try
                    {
                        s_processes = Process.GetProcessesByName("outlook");
                        if (s_processes.Length > 0)
                        {
                            EnumWindows(FoundWindowCallback, IntPtr.Zero); // Enumerate through all desktop windows
                        }
                        else
                        {
                            if (lastStartTry == null || lastStartTry.Elapsed > TimeSpan.FromSeconds(30))
                            {
                                lastStartTry = Stopwatch.StartNew();
                                Process.Start("outlook");
                            }
                        }
                    }
                    finally
                    {
                        if (s_processes != null)
                        {
                            foreach (var process in s_processes)
                            {
                                process.Dispose();
                            }
                        }
                    }
                    Thread.Sleep(1000);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error");
            }
        }


        static bool FoundWindowCallback(IntPtr hWnd, IntPtr lParam)
        {
            // Check if the window is a window of outlool
            int processId;
            GetWindowThreadProcessId(hWnd, out processId);
            if (s_processes.Any(p => p.Id == processId))
            {
                // Check if visibile
                if (IsWindowVisible(hWnd))
                {
                    // Check if Reminder
                    int size = GetWindowTextLength(hWnd) + 1;
                    StringBuilder titleBuilder = new StringBuilder(size);
                    int copied = GetWindowText(hWnd, titleBuilder, size);
                    titleBuilder.Length = copied;
                    string title = titleBuilder.ToString();
                    foreach (var searchedTitle in s_searchedTitles)
                    {
                        if (title.Contains(searchedTitle))
                        {
                            // Restore minimized window
                            ShowWindow(hWnd, SW_RESTORE);
                            // Make top most
                            SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                            // Center window in screen
                            RECT windowRect;
                            GetWindowRect(hWnd, out windowRect);
                            Rectangle screen = Screen.FromHandle(hWnd).Bounds;
                            Point pt = new Point(screen.Left + screen.Width / 2 - (windowRect.Right - windowRect.Left) / 2, screen.Top + screen.Height / 2 - (windowRect.Bottom - windowRect.Top) / 2);
                            SetWindowPos(hWnd, IntPtr.Zero, pt.X, pt.Y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_SHOWWINDOW);
                        }
                    }
                }
            }
            return true;
        }

    }
}
