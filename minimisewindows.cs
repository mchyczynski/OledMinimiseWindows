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
   private const string ScriptStateSetting = "OledMinimizerScriptState";
   private const string ScriptStateSweepSetting = "OledMinimizerScriptStateSweep";
   private const string MinimizedWindowsListSetting = "OledMinimizerMinimizedWindowsList";
   private const string SweptdWindowsListSetting = "OledMinimizerSweptWindowsList";
   private const string MousePositionXSetting = "MousePositionXSetting";
   private const string MousePositionYSetting = "MousePositionYSetting";
   private const string RestoredState = "0";
   private const string MinimizedState = "1";

   private static readonly uint RESOLUTION_4K_WIDTH = 3840;
   private static readonly uint RESOLUTION_4K_HEIGHT = 2160;

   private static readonly uint RESOLUTION_2K_WIDTH = 2560;
   private static readonly uint RESOLUTION_2K_HEIGHT = 1440;
   private static readonly uint MOUSE_RESTORE_THRESHOLD = 200;
   private static readonly int SWEEP_SNAP_THRESHOLD = 150;
   private static readonly double SWEEP_NO_RESIZE_THRESHOLD = 0.8;
   private static readonly int TASKBAR_HEIGHT = 40;

   public static string KEY_SHIFT = "16";
   public static string KEY_CTRL = "17";
   public static string KEY_ALT = "18";

   private static readonly bool enableMouseMove = true;
   private static readonly bool enableDebugPrints = true;
   private static readonly bool prioritizeMinimizeDefault = true;
   private static readonly bool keepRestoringDefault = true;
   private static readonly bool enableForceRestore = true;
   private static readonly bool enableFocusMode = true;
   private static readonly bool enableSweepMode = true;
   private static readonly bool enableSweepModeSnap = true;
   private static readonly bool debugPrintDoMinRestore = enableDebugPrints && false;
   private static readonly bool debugPrintStartStop = enableDebugPrints && false;
   private static readonly bool debugPrintFindMonitorId = enableDebugPrints && false;
   private static readonly bool debugPrintNoMonitorFound = enableDebugPrints && true;
   private static readonly bool debugWindowFiltering = enableDebugPrints && false;
   private static readonly bool debugPrintMoveCursor = enableDebugPrints && false;
   private static readonly bool debugPrintDecideMinRestore = enableDebugPrints && false;
   private static readonly bool debugPrintCountToMin = enableDebugPrints && false;
   private static readonly bool debugPrintForceRestoreKey = enableDebugPrints && false;
   private static readonly bool debugPrintFocusModeKey = enableDebugPrints && false;
   private static readonly bool debugPrintFocusMode = enableDebugPrints && false;
   private static readonly bool debugPrintSweepModeKey = enableDebugPrints && false;
   private static readonly bool debugPrintSweepMode = enableDebugPrints && false;
   private static readonly bool debugPrintSweepModeCalcPos = enableDebugPrints && false;


   private static List<string> classnameBlacklist = new List<string> {"DFTaskbar", "DFTitleBarWindow", "Shell_TrayWnd",
                                                                       "tooltips", "Shell_InputSwitchTopLevelWindow",
                                                                       "Windows.UI.Core.CoreWindow", "Progman", "SizeTipClass",
                                                                       "DF", "WorkerW", "SearchPane"};
   private static List<string> textBlacklist = new List<string> {"Program Manager", "Volume Mixer", "Snap Assist", "Greenshot capture form",
                                                                  "Battery Information", "Date and Time Information", "Network Connections",
                                                                  "Volume Control", "Start", "Search"};

   public static void Run(IntPtr windowHandle)
   {
      bool windowsAreVisible = CountWindowsToMinimize() > 0;
      bool windowsWereMinimized = WereWindowsMinimized();
      bool windowsWereSwept = WereWindowsSwept();
      bool forceRestore = ShouldForceRestore();

      if (forceRestore) // modifier key is pressed, never minimize
      {
         if (debugPrintDecideMinRestore) MessageBox.Show($"force restore, windows were" + (windowsWereMinimized ? "" : "NOT ") + " minimized previously");
         // try restoring only saved windows first if there were any
         int restoredCount = RestoreWindows(!windowsWereMinimized);
         // optionally restore all other minimized if saved windows were manually restored
         if ((restoredCount < 1) && ShouldKeepRestoring() && windowsWereMinimized)
         {
            restoredCount = RestoreWindows(true); // restore all windows (not only saved but minimalized manually)
            if (restoredCount < 1) MessageBox.Show($"There is no other windows to restore (todo remove)");
         }
      }
      else if (!windowsWereMinimized && windowsAreVisible)
      {
         // nothing was previously minimized but there are visible windows, just minimize them
         if (debugPrintDecideMinRestore) MessageBox.Show($"NOT windowsWereMinimized && windowsAreVisible");
         int minimizedCount = MinimizeWindows();
         if (minimizedCount < 1) MessageBox.Show($"ERROR no windows were minimized but should be!");
      }
      else if (!windowsWereMinimized && !windowsAreVisible)
      {
         // nothing was previously minimized, nothing is visible, force all windows that may have been manually minimized to restore
         if (debugPrintDecideMinRestore) MessageBox.Show($"NOT windowsWereMinimized && NOT windowsAreVisible");
         int restoredCount = RestoreWindows(true); // force restore all windows
         if (restoredCount < 1) MessageBox.Show($"No windows were restored but that may be ok if there was none at all (todo remove)"); // todo remove
      }
      else if (windowsWereMinimized && windowsAreVisible)
      {
         // there were windows minimized but there are also manually restored or new windows visible
         // decide if we should restore saved windows or minimize visible ones
         if (ShouldPrioritizeMinimize())
         {
            if (debugPrintDecideMinRestore) MessageBox.Show($"windowsWereMinimized && windowsAreVisible && ShouldPrioritizeMinimize");
            int minimizedCount = MinimizeWindows();
            if (minimizedCount < 1) MessageBox.Show($"ERROR no windows were minimized but should be (prio min)!");
         }
         else // don't prioritize minimizing new windows and restore saved ones
         {
            if (debugPrintDecideMinRestore) MessageBox.Show($"windowsWereMinimized && windowsAreVisible && NOT ShouldPrioritizeMinimize");
            int restoredCount = RestoreWindows(false); // try restoring only saved windows
            if ((restoredCount < 1) && ShouldKeepRestoring()) // optionally restore all other minimized if saved windows were manually restored
            {
               restoredCount = RestoreWindows(true); // restore all windows (not only saved but minimalized manually)
               if (restoredCount < 1) MessageBox.Show($"There is no other windows to restore (todo remove)");
            }
         }
      }
      else if (windowsWereMinimized && !windowsAreVisible)
      {
         // there were windows minimized and there are no new visible windows, just restore saved windows
         if (debugPrintDecideMinRestore) MessageBox.Show($"windowsWereMinimized && NOT windowsAreVisible");
         int restoredCount = RestoreWindows(false); // restore only saved windows
         if (restoredCount < 1) MessageBox.Show($"ERROR no windows were restored but should be");
      }
      else
      {
         MessageBox.Show($"ERROR else state not expected!");
      }
   }

   private static bool WereWindowsMinimized()
   {
      string setting = BFS.ScriptSettings.ReadValue(ScriptStateSetting);
      return !string.IsNullOrEmpty(setting) && (setting.Equals(MinimizedState, StringComparison.Ordinal));
   }

   private static bool WereWindowsSwept()
   {
      string setting = BFS.ScriptSettings.ReadValue(ScriptStateSweepSetting);
      return !string.IsNullOrEmpty(setting) && (setting.Equals(MinimizedState, StringComparison.Ordinal));
   }

   public static int MinimizeWindows()
   {
      if (debugPrintStartStop) MessageBox.Show("start MIN");
      // this will store the windows that we are minimizing so we can restore them later
      string minimizedWindows = "";

      // get monitor ID and bounds of OLED monitor
      uint monitorIdOled = GetOledMonitorID();
      Rectangle monitorBoundsOled = BFS.Monitor.GetMonitorBoundsByID(monitorIdOled);

      // get windows to be minimized
      IntPtr[] windowsToMinimize = GetFilteredVisibleWindows(monitorIdOled);

      // check if focus mode was requested
      bool focusMode = IsFocusModeRequested();

      // check if sweep mode was requested
      bool sweepMode = IsSweepModeRequested();

      // get monitor ID and bounds of sweep-target monitor
      uint monitorIdSweep = GetSweepTargetMonitorID();
      Rectangle monitorBoundsSweep = BFS.Monitor.GetMonitorBoundsByID(monitorIdSweep);

      // save handle to currently active window
      IntPtr activeWindowHandle = BFS.Window.GetFocusedWindow();
      if (debugPrintFocusMode) MessageBox.Show($"focus window found: {BFS.Window.GetText(activeWindowHandle)}");

      // sweep mode in reverse order (the same as restoring) compared to minimizing windows
      if (sweepMode)
      {
         Array.Reverse(windowsToMinimize);
      }

      int minimizedWindowsCount = 0;
      // loop through all the visible windows on the monitor
      foreach (IntPtr window in windowsToMinimize)
      {
         // if focus mode enabled skip focused window from list to minimize
         if (focusMode && window == activeWindowHandle)
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"skipping focus window {BFS.Window.GetText(window)}");
            minimizedWindowsCount += 1; // treat active window in focus mode as if it was minimized
            continue;
         }

         if (sweepMode)
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"sweeping window {BFS.Window.GetText(window)}");
            SweepWindow(window, monitorBoundsOled, monitorBoundsSweep);
         }
         else // normal minimizing
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"minimizing window {BFS.Window.GetText(window)}");
            WindowUtils.MinimizeWindow(window);
         }

         // use variables for both minimizing and sweeping
         minimizedWindowsCount += 1;
         // add the window to the list of minimized windows
         minimizedWindows += window.ToInt64().ToString() + "|";
      }

      // change focus and move mouse only when not in focus mode and not sweeping
      if (!focusMode && !sweepMode)
      {
         // it is a fix in order to enable alt-tabbing back to top minimized window
         // and being able to restore minimized windows from taskbar with only 1 mouse click
         WindowUtils.FocusOnDekstop();

         // hide mouse cursor to primary monitor (if feature is enabled and at least one window was minimized)
         if (ShoudlMoveMouse() && (minimizedWindowsCount > 0)) HandleMouseOut();
      }

      if (sweepMode)
      {
         // save the list of windows that were swept
         BFS.ScriptSettings.WriteValue(SweptdWindowsListSetting, minimizedWindows);

         // set the script sweep state to MinimizedState
         BFS.ScriptSettings.WriteValue(ScriptStateSweepSetting, MinimizedState);
      }
      else // normal minimizing
      {
         // save the list of windows that were minimized
         BFS.ScriptSettings.WriteValue(MinimizedWindowsListSetting, minimizedWindows);

         // set the script state to MinimizedState
         BFS.ScriptSettings.WriteValue(ScriptStateSetting, MinimizedState);
      }

      if (debugPrintStartStop) MessageBox.Show($"finished MIN (minimized {minimizedWindowsCount}/{windowsToMinimize.Length} windows)");

      return minimizedWindowsCount;
   }

   public static int RestoreWindows(bool forceRestoreAll)
   {
      if (debugPrintStartStop) MessageBox.Show("start RESTORE");

      // First restore mouse cursor position if enabled
      if (ShoudlMoveMouse()) HandleMouseBack();

      // get windows to be restored
      List<IntPtr> windowsToRestore = new List<IntPtr>();
      if (forceRestoreAll) // restore all windows on OLED monitor
      {
         // get monitor ID of OLED monitor (assumption it is the only 4k monitor in the system)
         windowsToRestore = GetFilteredMinimizedWindows(GetOledMonitorID()).ToList();
      }
      else // only restore windows previously minimized
      {
         // get the windows that we minimized previously
         string savedWindows = BFS.ScriptSettings.ReadValue(MinimizedWindowsListSetting);

         string[] windowsToRestoreStrings = savedWindows.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
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

      // check if focus mode was requested
      bool focusMode = IsFocusModeRequested();

      IntPtr[] windowsToPushOnTop = new IntPtr[] { };
      if (focusMode)
      {
         // save list of currently restored windows to later push them on top of restored ones
         windowsToPushOnTop = GetFilteredVisibleWindows(GetOledMonitorID());
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

      // if in focus mode, push windows that were active on top of restored ones
      if (focusMode)
      {
         Array.Reverse(windowsToPushOnTop);
         foreach (IntPtr windowHandle in windowsToPushOnTop)
         {
            if (debugPrintDoMinRestore) MessageBox.Show($"Pushing window {BFS.Window.GetText(new IntPtr(windowHandle))} on top (focus mode)");
            WindowUtils.PushToTop(windowHandle);
         }
      }

      // clear the windows that we saved
      BFS.ScriptSettings.WriteValue(MinimizedWindowsListSetting, string.Empty);

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

   public static void SweepWindow(IntPtr windowHandle, Rectangle boundsFrom, Rectangle boundsTo)
   {
      if (debugPrintSweepMode) MessageBox.Show($"SweepMonitor:\nboundsFrom: {boundsFrom},\nboundsTo {boundsTo}");

      Rectangle newPos = CalculateSweptWindowPos(WindowUtils.GetBounds(windowHandle), boundsFrom, boundsTo);
      WindowUtils.SetSizeAndLocation(windowHandle, newPos.X, newPos.Y, newPos.Width, newPos.Height);
      WindowUtils.PushToTop(windowHandle);

      // WindowUtils.MaximizeWindow(windowHandle); // todo
   }

   public static Rectangle CalculateSweptWindowPos(Rectangle boundsWindow, Rectangle boundsFrom, Rectangle boundsTo)
   {
      Rectangle newWindowBounds = new Rectangle();

      if ((boundsWindow.Width < SWEEP_NO_RESIZE_THRESHOLD * boundsTo.Width) ||
          (boundsWindow.Height < SWEEP_NO_RESIZE_THRESHOLD * (boundsTo.Height - TASKBAR_HEIGHT)))
      {
         // small window in at least one direction, don't change size for now
         newWindowBounds.Width = boundsWindow.Width;
         newWindowBounds.Height = boundsWindow.Height;

         // calculate how far is the windows from old monitor borders (ignore taskbar size because of autohide)
         int leftDistOld = boundsWindow.X - boundsFrom.X;
         int rightDistOld = (boundsFrom.X + boundsFrom.Width) - (boundsWindow.X + boundsWindow.Width);
         int topDistOld = boundsWindow.Y - boundsFrom.Y;
         int bottomDistOld = (boundsFrom.Y + boundsFrom.Height) - (boundsWindow.Y + boundsWindow.Height);

         // calculate left-right (horizontal) and top-bottom (vertical) ratios of window position
         double horizontalRatio = (double)leftDistOld / (leftDistOld + rightDistOld);
         double verticalRatio = (double)topDistOld / (topDistOld + bottomDistOld);

         // calculate how much free space there will be on new monitor outside of the window
         int horizontalSpace = boundsTo.Width - newWindowBounds.Width;
         int verticalSpace = boundsTo.Height - newWindowBounds.Height - TASKBAR_HEIGHT;

         if (horizontalSpace > 0)
         {
            // set position to preserve left-right ratio from old monitor
            newWindowBounds.X = boundsTo.X + (int)(horizontalRatio * horizontalSpace);
         }
         else
         {
            // there is no space left, max out window horizontally
            newWindowBounds.X = boundsTo.X;
            // override window witdth, it needs to be smaller
            newWindowBounds.Width = boundsTo.Width;

            // optionally snap top or bottom if window was resized horizontally
            if (enableSweepModeSnap)
            {
               int topDistNew = newWindowBounds.Y - boundsTo.Y;
               int bottomDistNew = (boundsTo.Y + boundsTo.Height - TASKBAR_HEIGHT) - (newWindowBounds.Y + newWindowBounds.Height);

               if (topDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap top");
                  newWindowBounds.Y = 0; // snap window to top, no size change
               }

               if (bottomDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap bottom");
                  newWindowBounds.Y = boundsTo.Height - TASKBAR_HEIGHT - newWindowBounds.Height; // snap window to bottom, no size change
               }

               if (debugPrintSweepModeCalcPos) MessageBox.Show($"topDistNew {topDistNew}\nbottomDistNew {bottomDistNew}\n\n" +
                                                $"newWindowBounds.Y {newWindowBounds.Y}");
            }

         }

         if (verticalSpace > 0)
         {
            // set position to preserve top-bottom ratio from old monitor
            newWindowBounds.Y = boundsTo.Y + (int)(verticalRatio * verticalSpace);
         }
         else
         {
            // if (debugPrintSweepModeCalcPos) MessageBox.Show($"Overriding vertical Y and Height");
            // there is no space left, max out window vertically
            newWindowBounds.Y = boundsTo.Y;
            // override window height, it needs to be smaller
            newWindowBounds.Height = boundsTo.Height - TASKBAR_HEIGHT;

            // optionally snap left or right if window was resized vertically
            if (enableSweepModeSnap)
            {
               int leftDistNew = newWindowBounds.X - boundsTo.X;
               int rightDistNew = (boundsTo.X + boundsTo.Width) - (newWindowBounds.X + newWindowBounds.Width);

               if (leftDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap left");
                  newWindowBounds.X = 0; // snap window to left, no size change
               }

               if (rightDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap right");
                  newWindowBounds.X = boundsTo.Width - newWindowBounds.Width; // snap window to right, no size change
               }

               if (debugPrintSweepModeCalcPos) MessageBox.Show($"leftDistNew {leftDistNew}\nrightDistNew {rightDistNew}\n\n" +
                                                               $"newWindowBounds.X {newWindowBounds.X}");
            }
         }


         if (debugPrintSweepModeCalcPos) MessageBox.Show($"\nboundsWindow\t{boundsWindow}\nboundsFrom\t{boundsFrom}\nboundsTo\t{boundsTo}\n\n" +
                         $"leftDistOld {leftDistOld}\nrightDistOld {rightDistOld}\ntopDistOld {topDistOld}\nbottomDistOld {bottomDistOld}\n\n" +
                         $"horizontalRatio {horizontalRatio}\nverticalRatio {verticalRatio}\n\n" +
                         $"horizontalSpace {horizontalSpace}\nverticalSpace {verticalSpace}\n\n" +
                         $"newWindowBounds.X {newWindowBounds.X}\nnewWindowBounds.Y {newWindowBounds.Y}\n\n" +
                         $"newWindowBounds.Width {newWindowBounds.Width}\nnewWindowBounds.Height {newWindowBounds.Height}");
      }
      else
      {
         // set size to target monitor size (excluding taskbar)
         newWindowBounds.X = boundsTo.X;
         newWindowBounds.Y = boundsTo.Y;
         newWindowBounds.Width = boundsTo.Width;
         newWindowBounds.Height = boundsTo.Height - TASKBAR_HEIGHT;

         if (debugPrintSweepModeCalcPos) MessageBox.Show($"\nboundsWindow\t{boundsWindow}\nboundsFrom\t{boundsFrom}\nboundsTo\t{boundsTo}\n\n" +
                         $"newWindowBounds.X {newWindowBounds.X}\nboundsTarget.Y {newWindowBounds.Y}\n\n" +
                         $"newWindowBounds.Width {newWindowBounds.Width}\nnewWindowBounds.Height {newWindowBounds.Height}");
      }

      return newWindowBounds;
   }

   public static uint GetOledMonitorID()
   {
      // assume OLED monitor is the first 4K monitor in the system
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

   public static uint GetSweepTargetMonitorID()
   {
      // assume we sweep windows to the first 2K monitor we find
      foreach (uint id in BFS.Monitor.GetMonitorIDs())
      {
         Rectangle bounds = BFS.Monitor.GetMonitorBoundsByID(id);
         if (bounds.Width == RESOLUTION_2K_WIDTH && bounds.Height == RESOLUTION_2K_HEIGHT)
         {
            if (debugPrintFindMonitorId) MessageBox.Show($"found 2k monitor with ID: {id}");
            return id;
         }
      }

      MessageBox.Show($"ERROR! did not find monitor with 2K resolution");
      return UInt32.MaxValue;
   }

   public static int CountWindowsToMinimize()
   {
      uint monitorId = GetOledMonitorID();
      IntPtr[] windowsToMinimize = GetFilteredVisibleWindows(monitorId);

      if (debugPrintCountToMin) MessageBox.Show($"CountWindowsToMinimize: {windowsToMinimize.Length}");
      foreach (IntPtr window in windowsToMinimize)
      {
         if (debugPrintCountToMin) MessageBox.Show($"To minimize: \n\n|{BFS.Window.GetText(window)}|\n\n|{BFS.Window.GetClass(window)}|");
      }

      return windowsToMinimize.Length;
   }

   public static bool ShouldPrioritizeMinimize()
   {
      // todo store in settings
      return prioritizeMinimizeDefault;
   }

   public static bool ShouldKeepRestoring()
   {
      // todo store in settings
      return keepRestoringDefault;
   }

   public static bool ShoudlMoveMouse()
   {
      return enableMouseMove && !IsFocusModeRequested();
   }

   public static bool ShouldForceRestore()
   {
      // bool keyPressed = BFS.Input.IsMouseDown("1;");
      bool keyPressed = BFS.Input.IsKeyDown(KEY_SHIFT);
      if (debugPrintForceRestoreKey) MessageBox.Show($"ForceRestore key is" + (keyPressed ? "" : " NOT") + " pressed");

      return keyPressed & enableForceRestore;
   }

   public static bool IsFocusModeRequested()
   {
      // bool keyPressed = BFS.Input.IsMouseDown("2;");
      bool keyPressed = BFS.Input.IsKeyDown(KEY_CTRL);
      if (debugPrintFocusModeKey) MessageBox.Show($"FocusMode key is" + (keyPressed ? "" : " NOT") + " pressed");

      return keyPressed & enableFocusMode;
   }

   public static bool IsSweepModeRequested()
   {
      // bool keyPressed = BFS.Input.IsMouseDown("2;");
      bool keyPressed = BFS.Input.IsKeyDown(KEY_ALT);
      if (debugPrintSweepModeKey) MessageBox.Show($"SweepMode key is" + (keyPressed ? "" : " NOT") + " pressed");

      return keyPressed & enableSweepMode;
   }

   public static IntPtr[] GetFilteredVisibleWindows(uint monitorId)
   {
      IntPtr[] allWindows = BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId);
      IntPtr[] filteredWindows = allWindows.Where(windowHandle =>
      {
         if (IsWindowBlacklisted(windowHandle))
         {
            // MessageBox.Show($"W ignored blacklisted:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
            return false;
         }
         else
         {
            // MessageBox.Show($"W NOT ignored blacklisted:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         }

         if (BFS.Window.IsMinimized(windowHandle))
         {
            // MessageBox.Show($"W ignored minimized:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
            return false;
         }
         else
         {
            // MessageBox.Show($"W NOT ignored minimized:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         }

         // MessageBox.Show($"W NOT ignored at all:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         return true;
      }).ToArray();

      return filteredWindows;
   }

   public static IntPtr[] GetFilteredMinimizedWindows(uint monitorId)
   {
      // get minimized windows from OLED 4k monitor
      IntPtr[] allWindows = BFS.Window.GetVisibleAndMinimizedWindowHandles().Where(windowHandle =>
      {
         // ignore windows that are not minimized
         if (!BFS.Window.IsMinimized(windowHandle)) return false;

         // find monitor size of minimized window
         Rectangle currentWindowMonitorBounds = WindowUtils.GetMonitorBoundsFromWindow(windowHandle);

         // filter window out when it would be restored to other monitors than OLED 4K
         if (currentWindowMonitorBounds.Width != RESOLUTION_4K_WIDTH ||
             currentWindowMonitorBounds.Height != RESOLUTION_4K_HEIGHT)
         {
            return false;
         }
         return true;
      }).ToArray();

      IntPtr[] filteredWindows = allWindows.Where(windowHandle => { return !IsWindowBlacklisted(windowHandle); }).ToArray();
      return filteredWindows;
   }

   public static bool IsWindowBlacklisted(IntPtr windowHandle)
   {
      // ignore windows based on classname blacklist
      string classname = BFS.Window.GetClass(windowHandle);
      if (classnameBlacklist.Exists(blacklistItem =>
            {
               if (classname.StartsWith(blacklistItem, StringComparison.Ordinal))
               {
                  if (debugWindowFiltering) MessageBox.Show($"Ignored bacause of class: {classname}|{blacklistItem}");
                  return true;
               }
               return false;
            })
         )
      {
         return true;
      }

      // ignore windows based on empty text
      string text = BFS.Window.GetText(windowHandle);
      if (string.IsNullOrEmpty(text))
      {
         if (debugWindowFiltering) MessageBox.Show($"Ignored bacause of empty text (classname: {classname})");
         return true;
      }

      // ignore windows based on text blacklist 
      if (textBlacklist.Exists(blacklistItem =>
            {
               if (text.Equals(blacklistItem, StringComparison.Ordinal))
               {
                  if (debugWindowFiltering) MessageBox.Show($"Ignored bacause of text: {text}|{blacklistItem}|");
                  return true;
               }
               return false;
            })
         )
      {
         return true;
      }

      // ignore windows with wrong size
      Rectangle windowRect = WindowUtils.GetBounds(windowHandle);
      if (windowRect.Width <= 0 || windowRect.Height <= 0)
      {
         // todo is it needed? remove or add if(debugWindowFiltering) 
         MessageBox.Show($"Filtered out windows wrong size (w{windowRect.Width}, h{windowRect.Height}). classname: {classname}, text: {text})");
         return true;
      }

      return false;
   }

   public static (int X, int Y) GetMouseHideTarget()
   {
      Rectangle bounds = BFS.Monitor.GetPrimaryMonitorBounds();
      return (bounds.Width / 2, bounds.Height / 2);
   }


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
      private static readonly int SW_SHOWMAXIMIZED = 3; // SW_MAXIMIZE
      private static readonly int SW_SHOWNOACTIVATE = 4;
      private static readonly int SW_SHOW = 5;
      private static readonly int SW_MINIMIZE = 6;
      private static readonly int SW_SHOWMINNOACTIVE = 7;
      private static readonly int SW_SHOWNA = 8;
      private static readonly int SW_RESTORE = 9;
      private static readonly int DWMWA_EXTENDED_FRAME_BOUNDS = 0x9;
      private static readonly uint MONITOR_DEFAULTTONEAREST = 2;
      private static readonly uint SWP_NOSIZE = 0x0001;
      private static readonly uint SWP_NOMOVE = 0x0002;
      private static readonly uint SWP_NOACTIVATE = 0x0010;
      private static readonly uint SWP_NOOWNERZORDER = 0x0200;
      private static readonly uint SWP_NOZORDER = 0x0004;
      private static readonly uint SWP_FRAMECHANGED = 0x0020;

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

      public static void MinimizeWindow(IntPtr windowHandle)
      {
         // ShowWindow(windowHandle, SW_MINIMIZE); // activates next window than currently minimized, pushes some windows at the end of alt-tab
         ShowWindow(windowHandle, SW_SHOWMINIMIZED); // leaves windows on top of alt-tab
         // ShowWindow(windowHandle, SW_SHOWMINNOACTIVE); // pushes windows to back of alt-tab
      }

      public static void MaximizeWindow(IntPtr windowHandle)
      {
         ShowWindow(windowHandle, SW_SHOWMAXIMIZED);
      }

      public static void RestoreWindow(IntPtr windowHandle)
      {
         ShowWindow(windowHandle, SW_RESTORE);
      }

      public static void FocusOnDekstop()
      {
         IntPtr hWndDesktop = GetShellWindow();
         if (hWndDesktop == IntPtr.Zero)
         {
            MessageBox.Show($"ERROR Desktop window not found!");
            return;
         }
         bool success = SetForegroundWindow(hWndDesktop);
         if (!success)
         {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"FocusOnDekstop error: {errorCode}");
         }
      }

      public static void PushToTop(IntPtr windowHandle)
      {
         if (debugPrintDoMinRestore) MessageBox.Show($"pushing\n|{BFS.Window.GetText(windowHandle)}|\non top");
         bool result = SetForegroundWindow(windowHandle);
         System.Threading.Thread.Sleep(30);
         if (!result)
         {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"SetForegroundWindow error: {errorCode}");
         }
      }
   } // WindowUtils
}