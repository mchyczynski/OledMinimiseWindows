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
   private const string MousePositionXSetting = "MousePositionXSetting";
   private const string MousePositionYSetting = "MousePositionYSetting";
   private const string RestoredState = "0";
   private const string MinimizedState = "1";

   private static readonly uint RESOLUTION_4K_WIDTH = 3840;
   private static readonly uint RESOLUTION_4K_HEIGHT = 2160;

   private static readonly uint RESOLUTION_2K_WIDTH = 2560;
   private static readonly uint RESOLUTION_2K_HEIGHT = 1440;
   private static readonly uint MOUSE_RESTORE_THRESHOLD = 200;

   private static bool enableMouseMove = true;
   private static bool enableDebugPrints = true;
   private static bool debugPrintDoMinRestore = enableDebugPrints && false;
   private static bool debugPrintStartStop = enableDebugPrints && false;
   private static bool debugPrintFindMonitorId = enableDebugPrints && false;
   private static bool debugPrintNoMonitorFound = enableDebugPrints && true;
   private static bool debugWindowFiltering = enableDebugPrints && false;
   private static bool debugPrintMoveCursor = enableDebugPrints && false;



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
         int minimizedCount = MinimizeWindows();

         // restore all windows when there was nothing to minimize
         if (minimizedCount < 1)
         {
            RestoreWindows(true); // force all windows restore
         }
      }
      else
      {
         int restoredCount = RestoreWindows(false);

         // minimize windows when there was nothing to restore
         if (restoredCount < 1)
         {
            MinimizeWindows();
         }
      }
   }

   // script should restore if there is no ScriptStateSetting set, or if the setting is equal to RestoredState
   // todo handle force minimize when there is at least one not minimized window
   private static bool ShouldMinimize()
   {
      //return true;
      string setting = BFS.ScriptSettings.ReadValue(ScriptStateSetting);
      return (setting.Length == 0) || (setting.Equals(RestoredState, StringComparison.Ordinal));
   }

   public static int MinimizeWindows()
   {
      if (debugPrintStartStop) MessageBox.Show("start MIN");
      // this will store the windows that we are minimizing so we can restore them later
      string minimizedWindows = "";

      // get monitor ID of OLED monitor (assumption it is the only 4k monitor in the system)
      uint monitorId = GetOledMonitorID();

      // get windows to be minimized
      IntPtr[] windowsToMinimize = GetFilteredVisibleWindows(monitorId);

      int minimizedWindowsCount = 0;
      // loop through all the visible windows on the monitor
      foreach (IntPtr window in windowsToMinimize)
      {
         // minimize the window
         if (!BFS.Window.IsMinimized(window))
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"minimizing window {BFS.Window.GetText(window)}");
            WindowUtils.MinimizeWindow(window);
            minimizedWindowsCount += 1;

            // add the window to the list of windows
            minimizedWindows += window.ToInt64().ToString() + "|";
         }
         else
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"already minimized window {BFS.Window.GetText(window)}");
         }

      }

      // hide mouse cursor to primary monitor (if feature is enabled and at least one window was minimized)
      if (enableMouseMove && (minimizedWindowsCount > 0)) HandleMouseOut();

      // save the list of windows that were minimized
      BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, minimizedWindows);

      // set the script state to MinimizedState
      BFS.ScriptSettings.WriteValue(ScriptStateSetting, MinimizedState);

      if (debugPrintStartStop) MessageBox.Show($"finished MIN (minimized {minimizedWindowsCount}/{windowsToMinimize.Length} windows)");

      return minimizedWindowsCount;
   }

   public static int RestoreWindows(bool forceRestoreAll)
   {
      if (debugPrintStartStop) MessageBox.Show("start RESTORE");

      // we are in the normalize window state
      // get the windows that we minimized previously
      string windows = BFS.ScriptSettings.ReadValue(MinimizedWindowsSetting);

      // First restore mouse cursor position if enabled
      if (enableMouseMove) HandleMouseBack();

      // get windows to be restored
      List<IntPtr> windowsToRestore = new List<IntPtr>();
      if (forceRestoreAll) // restore all windows on OLED monitor
      {
         // get monitor ID of OLED monitor (assumption it is the only 4k monitor in the system)
         windowsToRestore = GetFilteredMinimizedWindows(GetOledMonitorID()).ToList();
      }
      else // only restore windows previously minimized
      {
         string[] windowsToRestoreStrings = windows.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
         Array.Reverse(windowsToRestoreStrings);
         foreach (string window in windowsToRestoreStrings)
         {
            // try to turn the string into a long value
            // if we can't convert it, go to the next setting
            long windowHandleValue;
            if (!Int64.TryParse(window, out windowHandleValue))
               continue;

            windowsToRestore.Add(new IntPtr(windowHandleValue));
         }
      }

      int restoredWindowsCount = 0;
      // loop through each window to restore
      foreach (IntPtr windowHandle in windowsToRestore)
      {
         if (BFS.Window.IsMinimized(windowHandle))
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"restoring window {BFS.Window.GetText(new IntPtr(windowHandle))}");
            WindowUtils.RestoreWindow(windowHandle);
            WindowUtils.PushToTop(windowHandle);
            restoredWindowsCount += 1;
         }
         else
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"already restored window {BFS.Window.GetText(new IntPtr(windowHandle))}");
         }
      }

      // clear the windows that we saved
      BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, string.Empty);

      // set the script to RestoredState
      BFS.ScriptSettings.WriteValue(ScriptStateSetting, RestoredState);
      if (debugPrintStartStop) MessageBox.Show($"finished RESTORE (restored {restoredWindowsCount}/{windowsToRestore.Count} windows)");

      return restoredWindowsCount;
   }

   public static void HandleMouseOut()
   {
      // ccheck if mouse is on 4K OLED monitor so it should be moved
      int mouseX = BFS.Input.GetMousePositionX();
      int mouseY = BFS.Input.GetMousePositionY();

      // BFS.Monitor.GetMonitorBoundsByMouseCursor()
      Rectangle mouseMonitorBounds = BFS.Monitor.GetMonitorBoundsByXY(mouseX, mouseY);

      // mouse on monitor 4K OLED
      if (mouseMonitorBounds.Width == RESOLUTION_4K_WIDTH && mouseMonitorBounds.Height == RESOLUTION_4K_HEIGHT)
      {
         // store mouse position before moving it
         BFS.ScriptSettings.WriteValueInt(MousePositionXSetting, mouseX);
         BFS.ScriptSettings.WriteValueInt(MousePositionYSetting, mouseY);

         // move cursor to primary monitor
         var (mouseHideTargetX, mouseHideTargetY) = GetMouseHideTarget();
         BFS.Input.SetMousePosition(mouseHideTargetX, mouseHideTargetY);
         if (debugPrintMoveCursor) MessageBox.Show($"HandleMouseOut: hiding mouse from mouseOldX({mouseX}) mouseOldY({mouseY})");
      }
      else // mouse on other monitor
      {
         // clear stored mouse position
         BFS.ScriptSettings.DeleteValue(MousePositionXSetting);
         BFS.ScriptSettings.DeleteValue(MousePositionYSetting);
         if (debugPrintMoveCursor) MessageBox.Show($"HandleMouseOut: skip hiding mouse because not on 4k OLED monitor");
      }
   }

   public static void HandleMouseBack()
   {
      // read old mouse position
      int mouseOldX = BFS.ScriptSettings.ReadValueInt(MousePositionXSetting);
      int mouseOldY = BFS.ScriptSettings.ReadValueInt(MousePositionYSetting);

      // abort when no stored position
      if (mouseOldX == 0 && mouseOldY == 0)
      {
         // ignore fact that mouse may actually be saved at 0,0 pos as a minor problem
         return;
      }

      // read currect mouse position
      int mouseX = BFS.Input.GetMousePositionX();
      int mouseY = BFS.Input.GetMousePositionY();

      // check if mouse was moved after hiding it from 4k OLED monitor 
      // which is to check if current position differs enough from mouse hide position
      var (mouseHideTargetX, mouseHideTargetY) = GetMouseHideTarget();
      int diffX = Math.Abs(mouseX - mouseHideTargetX);
      int diffY = Math.Abs(mouseY - mouseHideTargetY);

      bool wasMoved = diffX > MOUSE_RESTORE_THRESHOLD || diffY > MOUSE_RESTORE_THRESHOLD;
      if (!wasMoved)
      {
         // restore mouse position because it wasn't moved enough
         if (debugPrintMoveCursor) MessageBox.Show($"HandleMouseBack: restoring to mouseOldX({mouseOldX}) mouseOldY({mouseOldY})");
         BFS.Input.SetMousePosition(mouseOldX, mouseOldY);
      }
      else
      {
         if (debugPrintMoveCursor) MessageBox.Show($"HandleMouseBack: skiping restoring mouse was moved too much:\n" +
                                                    $"mouseOldX({mouseOldX}) mouseOldY({mouseOldY})\n" +
                                                    $"diffX({diffX}) diffY({diffY})\n" +
                                                    $"mouseX({mouseX}) mouseY({mouseY})\n" +
                                                    $"hideX({mouseHideTargetX}) hideY({mouseHideTargetY})");
      }

      // clear stored mouse position
      BFS.ScriptSettings.DeleteValue(MousePositionXSetting);
      BFS.ScriptSettings.DeleteValue(MousePositionYSetting);
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

      MessageBox.Show($"ERROR! did not find monitor with 4K resolution");
      return UInt32.MaxValue;
   }

   public static IntPtr[] GetFilteredVisibleWindows(uint monitorId)
   {
      IntPtr[] allWindows = BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId);
      IntPtr[] filteredWindows = allWindows.Where(FilterBlacklistedWindowsOut).ToArray();

      return filteredWindows;
   }

   public static IntPtr[] GetFilteredMinimizedWindows(uint monitorId)
   {
      // get minimized windows from OLED 4k monitor
      IntPtr[] allWindows = BFS.Window.GetVisibleAndMinimizedWindowHandles().Where(windowHandle =>
      {
         // ignore windows that are not minimized
         if (!BFS.Window.IsMinimized(windowHandle)) return false;

         // find monitor size od minimized window
         Rectangle currentWindowMonitorBounds = WindowUtils.GetMonitorBoundsFromWindow(windowHandle);

         // filter window out when it would be restored to other monitors than OLED 4K
         if (currentWindowMonitorBounds.Width != RESOLUTION_4K_WIDTH ||
             currentWindowMonitorBounds.Height != RESOLUTION_4K_HEIGHT)
         {
            return false;
         }
         return true;
      }).ToArray();

      IntPtr[] filteredWindows = allWindows.Where(FilterBlacklistedWindowsOut).ToArray();
      return filteredWindows;
   }

   public static bool FilterBlacklistedWindowsOut(IntPtr windowHandle)
   {
      // ignore windows based on classname blacklist
      string classname = BFS.Window.GetClass(windowHandle);
      if (classnameBlacklist.Exists(blacklistItem =>
      {
         if (classname.StartsWith(blacklistItem, StringComparison.Ordinal))
         {
            //MessageBox.Show($"Ignored bacause of class: {classname}|{blacklistItem}");
            return true;
         }
         return false;
      }))
      {
         return false;
      }

      // ignore windows based on empty text
      string text = BFS.Window.GetText(windowHandle);
      if (string.IsNullOrEmpty(text))
      {
         // MessageBox.Show($"Ignored bacause of empty text");
         return false;
      }

      // ignore windows based on text blacklist 
      if (textBlacklist.Exists(blacklistItem =>
      {
         if (text.Equals(blacklistItem, StringComparison.Ordinal))
         {
            MessageBox.Show($"Ignored bacause of text: {text}|{blacklistItem}|");
            return true;
         }
         return false;
      }))
      {
         return false;
      }

      // ignore windows with wrong size
      Rectangle windowRect = WindowUtils.GetBounds(windowHandle);
      if (windowRect.Width <= 0 || windowRect.Height <= 0)
      {
         // todo is it needed? add if for print-debug-flag
         MessageBox.Show($"Filtered out windows wrong size (w{windowRect.Width}, h{windowRect.Height}). classname: {classname}, text: {text})");
         return false;
      }

      return true;
   }

   public static (int X, int Y) GetMouseHideTarget()
   {
      Rectangle bounds = BFS.Monitor.GetPrimaryMonitorBounds();
      return (bounds.Width / 2, bounds.Height / 2);
   }


   public static class WindowUtils
   {
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

      private static readonly int SW_SHOWNORMAL = 1;
      private static readonly int SW_SHOWMINIMIZED = 2;
      private static readonly int SW_SHOWNOACTIVATE = 4;
      private static readonly int SW_SHOW = 5;
      private static readonly int SW_MINIMIZE = 6;
      private static readonly int SW_SHOWMINNOACTIVE = 7;
      private static readonly int SW_SHOWNA = 8;
      private static readonly int SW_RESTORE = 9;
      private static readonly int DWMWA_EXTENDED_FRAME_BOUNDS = 0x9;
      private static readonly uint MONITOR_DEFAULTTONEAREST = 2;
      
      private static bool GetRectangleExcludingShadow(IntPtr handle, out RECT rect)
      {
         var result = DwmGetWindowAttribute(handle, DWMWA_EXTENDED_FRAME_BOUNDS, out rect, Marshal.SizeOf(typeof(RECT)));

         // remove additional semi-transparent pixels on borders to minimize gaps between windows
         rect.Left++;
         rect.Top++;
         rect.Bottom--;

         return result >= 0;
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

      public static Rectangle GetMonitorBoundsFromWindow(IntPtr windowHandle)
      {
         IntPtr monitorHandle = MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST);

         if (monitorHandle != IntPtr.Zero)
         {
            Rectangle bounds = GetMonitorBounds(monitorHandle);
            // MessageBox.Show($"Monitor Bounds: Left={bounds.Left}, Top={bounds.Top}, Right={bounds.Right}, Bottom={bounds.Bottom}");
            return bounds;
         }
         else
         {
            MessageBox.Show("Error in GetMonitorBoundsFromWindow! Failed to determine the monitor.");
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
            MessageBox.Show("ERROR in GetMonitorBounds! Failed to retrieve monitor information.");
         }
         return default;
      }

      public static void MinimizeWindow(IntPtr windowHandle)
      {
         // ShowWindow(windowHandle, SW_MINIMIZE); // activates next window than currently minimized
         // ShowWindow(windowHandle, SW_SHOWMINIMIZED); // activates currently minimized window (doesn't work?)
         ShowWindow(windowHandle, SW_SHOWMINNOACTIVE); // doesn't activate any window
      }
      public static void RestoreWindow(IntPtr windowHandle)
      {
         ShowWindow(windowHandle, SW_RESTORE);
      }

      public static void PushToTop(IntPtr windowHandle)
      {
         if (debugPrintDoMinRestore) MessageBox.Show($"pushing\n|{BFS.Window.GetText(windowHandle)}|\non top");
         bool result = SetForegroundWindow(windowHandle);
         if (!result)
         {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"SetForegroundWindow error: {errorCode}");
         }
      }
   } // WindowUtils
}