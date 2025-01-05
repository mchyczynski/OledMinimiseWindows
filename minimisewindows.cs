using System;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;

using System;
using System.Text;
using System.Reflection;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class DisplayFusionFunction
{
    private const string ScriptStateSetting = "CursorMonitorScriptState";
    private const string MinimizedWindowsSetting = "CursorMonitorMinimizedWindows";
    private const string MinimizedState = "0";
    private const string NormalizeState = "1";

    private const uint RESOLUTION_4K_WIDTH = 3840;
    private const uint RESOLUTION_4K_HEIGHT = 2160;

    private static bool enableDebugPrints = true;
    private static bool debugPrint = false;
    private static bool debugPrintStartStop = false;
    private static bool debugPrintFindMonitorId = enableDebugPrints && false;



    private static List<string> classnameBlacklist = new List<string> {"DFTaskbar", "DFTitleBarWindow", "Shell_TrayWnd",
                                                                       "tooltips", "Shell_InputSwitchTopLevelWindow",
                                                                       "Windows.UI.Core.CoreWindow", "Progman", "SizeTipClass",
                                                                       "DF", "WorkerW", "SearchPane"};
    private static List<string> textBlacklist = new List<string> {"Program Manager", "Volume Mixer", "Snap Assist", "Greenshot capture form",
                                                                  "Battery Information", "Date and Time Information", "Network Connections",
                                                                  "Volume Control", "Start", "Search"};

    public static void Run(IntPtr windowHandle)
    {
        if (ShouldMinimize())
        {
            MinimizeWindows();
        }
        else
        {
            MaximizeWindows();
        }
    }

    // script is in minimize state if there is no setting, or if the setting is equal to MinimizedState
    private static bool ShouldMinimize()
    {
        //return true;
        string setting = BFS.ScriptSettings.ReadValue(ScriptStateSetting);
        return (setting.Length == 0) || (setting.Equals(MinimizedState, StringComparison.Ordinal));
    }

    public static void MinimizeWindows()
    {
        if (debugPrintStartStop) MessageBox.Show("start MIN");
        // this will store the windows that we are minimizing so we can restore them later
        string minimizedWindows = "";

        // get monitor ID of OLED monitor (assumption it is the only 4k monitor in the system)
        uint monitorId = GetOledMonitorID();

        // loop through all the visible windows on the cursor monitor
        foreach (IntPtr window in GetFilteredWindows(monitorId))
        {
            // minimize the window
            if (debugPrint) MessageBox.Show($"minimizing {BFS.Window.GetText(window)}");
            BFS.Window.Minimize(window);

            // add the window to the list of windows
            minimizedWindows += window.ToInt64().ToString() + "|";
        }

        // save the list of windows we minimized
        BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, minimizedWindows);

        // set the script state to NormalizeState
        BFS.ScriptSettings.WriteValue(ScriptStateSetting, NormalizeState);

        if (debugPrintStartStop) MessageBox.Show("finished MIN");
    }

    public static void MaximizeWindows()
    {
        if (debugPrintStartStop) MessageBox.Show("start MAX");

        // we are in the normalize window state
        // get the windows that we minimized previously
        string windows = BFS.ScriptSettings.ReadValue(MinimizedWindowsSetting);

        // loop through each setting
        foreach (string window in windows.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
        {
            // try to turn the string into a long value
            // if we can't convert it, go to the next setting
            long windowHandleValue;
            if (!Int64.TryParse(window, out windowHandleValue))
                continue;

            // restore the window
            BFS.Window.Restore(new IntPtr(windowHandleValue));
        }

        // clear the windows that we saved
        BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, string.Empty);

        // set the script to MinimizedState
        BFS.ScriptSettings.WriteValue(ScriptStateSetting, MinimizedState);
        if (debugPrintStartStop) MessageBox.Show("finished MAX");
    }

    public static uint GetOledMonitorID()
    {
        foreach (uint id in BFS.Monitor.GetMonitorIDs())
        {
            Rectangle bounds = BFS.Monitor.GetMonitorBoundsByID(id);
            if (bounds.Width == RESOLUTION_4K_WIDTH && bounds.Height == RESOLUTION_4K_HEIGHT)
            {
                if (debugPrintFindMonitorId) MessageBox.Show($"found 4k monitor with ID: {id}");
                return id;
            }
        }

        MessageBox.Show($"ERROR! did not find monitor with 4k resolution");
        return UInt32.MaxValue;
    }

    public static IntPtr[] GetFilteredWindows(uint monitorId)
    {
        IntPtr[] allWindows = BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId);

        IntPtr[] filteredWindows = allWindows.Where(windowHandle =>
        {

            // Ignore already minimized windows
            if (BFS.Window.IsMinimized(windowHandle))
            {
                return false;
            }

            // Ignore windows based on classname blacklist
            string classname = BFS.Window.GetClass(windowHandle);
            if (classnameBlacklist.Exists(blacklistItem => classname.StartsWith(blacklistItem, StringComparison.Ordinal)))
            {
                return false;
            }

            // Ignore windows based on text blacklist or is empty
            string text = BFS.Window.GetText(windowHandle);
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            if (textBlacklist.Exists(blacklistItem => text.StartsWith(blacklistItem, StringComparison.Ordinal)))
            {
                return false;
            }

            // Ignore windows with wrong size
            Rectangle windowRect = WindowUtils.GetBounds(windowHandle);
            if (windowRect.Width <= 0 || windowRect.Height <= 0)
            {
                MessageBox.Show($"Filtered out windows wrong size (w{windowRect.Width}, h{windowRect.Height}. classname: {classname}, text: {text})");
                return false;
            }

            return true;
        }).ToArray();

        return filteredWindows;
    }


    public static class WindowUtils
    {
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags); // including shadow

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect); // including shadow

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport(@"dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

        [Serializable, StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public Rectangle ToRectangle()
            {
                return Rectangle.FromLTRB(Left, Top, Right, Bottom);
            }
            public override string ToString()
            {
                return $"Left.{Left} Right.{Right} Top.{Top} Bottom.{Bottom} // " + ToRectangle().ToString();
            }
        }

        private static readonly IntPtr HWND_TOP = new IntPtr(0);
        private static readonly uint SWP_NOSIZE = 0x0001;
        private static readonly uint SWP_NOMOVE = 0x0002;
        private static readonly uint SWP_NOACTIVATE = 0x0010;
        private static readonly uint SWP_NOOWNERZORDER = 0x0200;
        private static readonly uint SWP_NOZORDER = 0x0004;
        private static readonly uint SWP_FRAMECHANGED = 0x0020;
        private static readonly uint RDW_INVALIDATE = 0x0001;
        private static readonly uint RDW_ALLCHILDREN = 0x0080;
        private static readonly uint RDW_UPDATENOW = 0x0100;
        private static readonly int DWMWA_EXTENDED_FRAME_BOUNDS = 0x9;


        private static bool GetRectangleExcludingShadow(IntPtr handle, out RECT rect)
        {
            var result = DwmGetWindowAttribute(handle, DWMWA_EXTENDED_FRAME_BOUNDS, out rect, Marshal.SizeOf(typeof(RECT)));

            // remove additional semi-transparent pixels on borders to minimize gaps between windows
            rect.Left++;
            rect.Top++;
            rect.Bottom--;

            return result >= 0;
        }
        private static Rectangle CompensateForShadow(IntPtr windowHandle, int x, int y, int w, int h)
        {
            RECT excludeShadow = new RECT();
            RECT includeShadow = new RECT();

            if (!GetRectangleExcludingShadow(windowHandle, out excludeShadow))
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                MessageBox.Show($"ERROR CompensateForShadow-GetWindowRectangle windows API: {errorCode}\n\n" +
                                $"text: |{text}|\n\n" +
                                $"requested pos: x.{x} y.{y} w.{w} h.{h}");
            }

            if (!GetWindowRect(windowHandle, out includeShadow)) // including shadow
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                MessageBox.Show($"ERROR CompensateForShadow-GetWindowRect windows API: {errorCode}\n\n" +
                                $"text: |{text}|\n\n" +
                                $"requested pos: x.{x} y.{y} w.{w} h.{h}");
            }

            RECT shadow = new RECT();
            shadow.Left = Math.Abs(includeShadow.Left - excludeShadow.Left);//+1;
            shadow.Right = Math.Abs(includeShadow.Right - excludeShadow.Right);
            shadow.Top = Math.Abs(includeShadow.Top - excludeShadow.Top);// +1;
            shadow.Bottom = Math.Abs(includeShadow.Bottom - excludeShadow.Bottom);//+1;

            // compensate requested x, y, width and height with shadow
            Rectangle result = new Rectangle(
                x - shadow.Left, // windowX
                y - shadow.Top, // windowY
                w + shadow.Right + shadow.Left, // windowWidth
                h + shadow.Bottom + shadow.Top // windowHeight
            );

            return result;
        }

        public static void SetLocation(IntPtr windowHandle, int x, int y)
        {
            SetSizeAndLocation(windowHandle, x, y, 0, 0);
        }

        public static void SetSizeAndLocation(IntPtr windowHandle, int x, int y, int w, int h)
        {
            Rectangle newPos = CompensateForShadow(windowHandle, x, y, w, h);

            uint flags = SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_FRAMECHANGED;

            if (w == 0 && h == 0)
            {
                flags = flags | SWP_NOSIZE;
            }

            if (!SetWindowPos(windowHandle, windowHandle, newPos.X, newPos.Y, newPos.Width, newPos.Height, flags))
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                MessageBox.Show($"ERROR SetSizeAndLocation-SetWindowPos windows API: {errorCode}\n\n" +
                                $"text: |{text}|\n\n" +
                                $"requested pos: x.{x} y.{y} w.{w} h.{h}");
            }
        }
        public static Rectangle GetBounds(IntPtr windowHandle)
        {
            int windowX = 0, windowY = 0, windowWidth = 0, windowHeight = 0;
            RECT windowRect;

            if (GetRectangleExcludingShadow(windowHandle, out windowRect))
            {
                windowX = windowRect.Left;
                windowY = windowRect.Top;
                windowWidth = windowRect.Right - windowRect.Left;
                windowHeight = windowRect.Bottom - windowRect.Top;
            }
            else
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                MessageBox.Show($"ERROR GetBounds-GetWindowRect windows API: {errorCode}\n\ntext: |{text}|");
            }

            Rectangle rect = new Rectangle(
                windowX,
                windowY,
                windowWidth,
                windowHeight
            );
            return rect;
        }

        public static void RedrawWindow(IntPtr windowHandle)
        {
            bool result = RedrawWindow(windowHandle, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);

            if (!result)
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                Rectangle windowRectangle = WindowUtils.GetBounds(windowHandle);

                MessageBox.Show($"ERROR RedrawWindow windows API: {errorCode}\n\ntext: |{text}|\n\nrect: {windowRectangle.ToString()}");
            }
        }

        public static void ForceUpdateWindow(IntPtr windowHandle)
        {
            bool result = UpdateWindow(windowHandle);

            if (!result)
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                Rectangle windowRectangle = WindowUtils.GetBounds(windowHandle);

                MessageBox.Show($"ERROR UpdateWindow windows API: {errorCode}\n\ntext: |{text}|\n\nrect: {windowRectangle.ToString()}");
            }
        }

        public static void InvalidateRectangle(IntPtr windowHandle)
        {
            bool result = InvalidateRect(windowHandle, IntPtr.Zero, true);

            if (!result)
            {
                int errorCode = Marshal.GetLastWin32Error();
                string text = BFS.Window.GetText(windowHandle);
                Rectangle windowRectangle = WindowUtils.GetBounds(windowHandle);

                MessageBox.Show($"ERROR InvalidateRectangle windows API: {errorCode}\n\ntext: |{text}|\n\nrect: {windowRectangle.ToString()}");
            }
        }
    } // WindowUtils
}