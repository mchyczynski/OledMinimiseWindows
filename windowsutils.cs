using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

public static class WindowUtils
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetShellWindow();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport(@"dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags); // including shadow

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect); // including shadow

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

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

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public uint cbSize;
        public RECT rcMonitor; // The monitor bounds
        public RECT rcWork;    // The work area
        public uint dwFlags;   // Monitor flags
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public POINT minPosition;
        public POINT maxPosition;
        public RECT normalPosition;
    }

    private const int SW_SHOWNORMAL = 1;
    private const int SW_SHOWMINIMIZED = 2;
    private const int SW_SHOWMAXIMIZED = 3; // SW_MAXIMIZE
    private const int SW_SHOWNOACTIVATE = 4;
    private const int SW_SHOW = 5;
    private const int SW_MINIMIZE = 6;
    private const int SW_SHOWMINNOACTIVE = 7;
    private const int SW_SHOWNA = 8;
    private const int SW_RESTORE = 9;
    private const int DWMWA_EXTENDED_FRAME_BOUNDS = 0x9;
    private const uint MONITOR_DEFAULTTONEAREST = 2;
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_NOOWNERZORDER = 0x0200;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_FRAMECHANGED = 0x0020;

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
            MessageBox.Show($"ERROR <min> CompensateForShadow-GetWindowRectangle windows API: {errorCode}\n\n" +
                        $"windowHandle: |{windowHandle}|\n\n" +
                        $"text: |{text}|\n\n" +
                        $"requested pos: x.{x} y.{y} w.{w} h.{h}");
        }

        if (!GetWindowRect(windowHandle, out includeShadow)) // including shadow
        {
            int errorCode = Marshal.GetLastWin32Error();
            string text = BFS.Window.GetText(windowHandle);
            MessageBox.Show($"ERROR <min> CompensateForShadow-GetWindowRect windows API: {errorCode}\n\n" +
                        $"windowHandle: |{windowHandle}|\n\n" +
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

    public static Rectangle GetBounds(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            Log.E($"window not valid", new { windowHandle });
            return new Rectangle { };
        }

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
            MessageBox.Show($"ERROR <min> GetBounds-GetWindowRect windows API: {errorCode}\n\ntext: |{text}|");
        }

        Rectangle rect = new Rectangle(
            windowX,
            windowY,
            windowWidth,
            windowHeight
        );
        return rect;
    }

    public static Rectangle GetRestoredBounds(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in GetRestoredBounds: {windowHandle}");
            return new Rectangle { };
        }

        WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
        placement.length = Marshal.SizeOf(typeof(WINDOWPLACEMENT));

        if (GetWindowPlacement(windowHandle, ref placement))
        {
            // The normalPosition field contains the restore bounds.
            RECT r = placement.normalPosition;
            return new Rectangle(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
        }
        else
        {
            int errorCode = Marshal.GetLastWin32Error();
            string text = BFS.Window.GetText(windowHandle);
            MessageBox.Show($"ERROR <min> GetRestoredBoundss-GetWindowPlacement windows API: {errorCode}\n\ntext: |{text}|");
        }
        return new Rectangle { };
    }

    public static Rectangle GetMonitorBoundsFromWindow(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in GetMonitorBoundsFromWindow: {windowHandle}");
            return new Rectangle { };
        }

        IntPtr monitorHandle = MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST);

        if (monitorHandle != IntPtr.Zero)
        {
            Rectangle bounds = GetMonitorBounds(monitorHandle);
            // MessageBox.Show($"Monitor Bounds: Left={bounds.Left}, Top={bounds.Top}, Right={bounds.Right}, Bottom={bounds.Bottom}");
            return bounds;
        }
        else
        {
            MessageBox.Show("ERROR <min> in GetMonitorBoundsFromWindow! Failed to determine the monitor.");
        }
        return default;
    }

    private static Rectangle GetMonitorBounds(IntPtr monitorHandle)
    {
        MONITORINFO monitorInfo = new MONITORINFO();
        monitorInfo.cbSize = (uint)Marshal.SizeOf(typeof(MONITORINFO));

        if (GetMonitorInfo(monitorHandle, ref monitorInfo))
        {
            return monitorInfo.rcMonitor.ToRectangle();
        }
        else
        {
            MessageBox.Show("ERROR <min> in GetMonitorBounds! Failed to retrieve monitor information.");
        }
        return default;
    }

    public static void SetSizeAndLocation(IntPtr windowHandle, Rectangle bounds)
    {
        SetSizeAndLocation(windowHandle, bounds.X, bounds.Y, bounds.Width, bounds.Height);
    }
    public static void SetSizeAndLocation(IntPtr windowHandle, int x, int y, int w, int h)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in SetSizeAndLocation: {windowHandle}");
            return;
        }

        Rectangle currentPos = GetBounds(windowHandle);
        Rectangle requestedPos = new Rectangle(x, y, w, h);
        Rectangle checkedPos = new Rectangle { };
        string text = BFS.Window.GetText(windowHandle);

        Rectangle newPos = CompensateForShadow(windowHandle, x, y, w, h);

        Func<string> debugInfo = () =>
                         $"text: |{text}|\n\n" +
                         $"current pos: {currentPos}\n" +
                         $"requested pos: {requestedPos}\n" +
                         $"shad-comp pos: {newPos}\n" +
                         $"checked pos: {checkedPos}\n\n";

        uint flags = SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_FRAMECHANGED;

        if (w == 0 && h == 0) flags = flags | SWP_NOSIZE;

        if (!SetWindowPos(windowHandle, windowHandle, newPos.X, newPos.Y, newPos.Width, newPos.Height, flags))
        {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"ERROR <min> SetSizeAndLocation-1st SetWindowPos windows API: {errorCode}\n\n" + debugInfo());
        }

        checkedPos = WindowUtils.GetBounds(windowHandle);

        if (checkedPos != requestedPos)
        {
            MessageBox.Show($"ERROR <min> SetSizeAndLocation-SetWindowPos wrong checked pos:\n\n" + debugInfo() + "moving again");

            if (!SetWindowPos(windowHandle, windowHandle, newPos.X, newPos.Y, newPos.Width, newPos.Height, flags))
            {
                int errorCode = Marshal.GetLastWin32Error();
                MessageBox.Show($"ERROR <min> SetSizeAndLocation-2nd SetWindowPos windows API: {errorCode}\n\n" + debugInfo());
            }

            checkedPos = WindowUtils.GetBounds(windowHandle);
            if (checkedPos != requestedPos)
                MessageBox.Show($"posistion sitll wrong:\n\n" + debugInfo());
            else
                MessageBox.Show($"posistion now OK\n\n" + debugInfo());
        }
    }

    public static void MinimizeWindow(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in MinimizeWindow: {windowHandle}");
            return;
        }
        // ShowWindow(windowHandle, SW_MINIMIZE); // activates next window than currently minimized, pushes some windows at the end of alt-tab
        ShowWindow(windowHandle, SW_SHOWMINIMIZED); // leaves windows on top of alt-tab
                                                    // ShowWindow(windowHandle, SW_SHOWMINNOACTIVE); // pushes windows to back of alt-tab
    }

    public static void MaximizeWindow(IntPtr windowHandle, int delay = 0)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in MaximizeWindow: {windowHandle}");
            return;
        }

        if (delay > 0) System.Threading.Thread.Sleep(delay);

        ShowWindow(windowHandle, SW_SHOWMAXIMIZED);
    }

    public static void RestoreWindow(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in RestoreWindow: {windowHandle}");
            return;
        }

        ShowWindow(windowHandle, SW_RESTORE);
    }

    public static void FocusOnDekstop()
    {
        IntPtr hWndDesktop = GetShellWindow();
        if (hWndDesktop == IntPtr.Zero)
        {
            MessageBox.Show($"ERROR <min> Desktop window not found!");
            return;
        }
        FocusOnWindow(hWndDesktop);
    }

    public static void FocusOnWindow(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            MessageBox.Show($"ERROR <min> window not valid in FocusOnWindow: {windowHandle}");
            return;
        }

        bool success = SetForegroundWindow(windowHandle);
        if (!success)
        {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"FocusOnWindow error: {errorCode}");
        }
    }

    public static void PushToTop(IntPtr windowHandle)
    {
        if (!IsWindow(windowHandle))
        {
            Log.E($"window not valid", new { windowHandle });
            return;
        }

        Log.D($"pushing window on top", () => new { text = BFS.Window.GetText(windowHandle) });
        bool result = SetForegroundWindow(windowHandle);
        System.Threading.Thread.Sleep(30);
        if (!result)
        {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"SetForegroundWindow error: {errorCode}");
        }
    }

    public static bool IsWindowValid(IntPtr windowHandle)
    {
        return IsWindow(windowHandle);
    }
} // WindowUtils
