using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class DisplayFusionFunction
{
    public static void Run(IntPtr windowHandle)
    {

        IntPtr[] allWindows = BFS.Window.GetVisibleAndMinimizedWindowHandles();
        foreach (IntPtr window in allWindows)
        {
            WindowUtils.StopFlashingNotification(window);
        }
        MessageBox.Show($"Stopped flashing windows");

    }

    public static class WindowUtils
    {
        // To support flashing.
        [DllImport("user32.dll", CallingConvention = CallingConvention.Cdecl)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);


        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        //Flash both the window caption and taskbar button.
        //This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags. 
        private const uint FLASHW_ALL = 3;

        // Stop flashing. The system restores the window to its original state.
        private const uint FLASHW_STOP = 0;

        // Flash continuously until the window comes to the foreground. 
        private const uint FLASHW_TIMERNOFG = 12;

        // Definicje flag dla SetWindowPos
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_SHOWWINDOW = 0x0040;


        // Performs no operation. An application sends the WM_NULL message if it wants to post a message that the recipient window will ignore.
        private const int WM_NULL = 0x0006;
        // Message 
        private const int WM_ACTIVATE = 0x0006;
        // Activated by some method other than a mouse click
        private const int WA_ACTIVE = 1;

        public static bool StopFlashingNotification(IntPtr windowHandle)
        {
            FLASHWINFO fInfo = new FLASHWINFO();

            fInfo.cbSize = Convert.ToUInt32(Marshal.SizeOf(fInfo));
            fInfo.hwnd = windowHandle;
            fInfo.dwFlags = FLASHW_STOP;
            fInfo.uCount = 0;
            fInfo.dwTimeout = 0;

            bool ret = FlashWindowEx(ref fInfo);

            // SendMessage(windowHandle, WM_ACTIVATE, (IntPtr)WA_ACTIVE, IntPtr.Zero);
            // SendMessage(windowHandle, WM_NULL, IntPtr.Zero, IntPtr.Zero);

            SetWindowPos(windowHandle, IntPtr.Zero, 0, 0, 0, 0, 
                SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            return ret;
        }

    } // WindowUtils
}