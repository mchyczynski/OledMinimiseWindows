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
   private const string ScriptStateMinSetting = "OledMinimizerScriptState";
   private const string ScriptStateSweepSetting = "OledMinimizerScriptStateSweep";
   private const string MinimizedWindowsListSetting = "OledMinimizerMinimizedWindowsList";
   private const string SweptdWindowsListSetting = "OledMinimizerSweptWindowsList";
   private const string MousePositionXSetting = "MousePositionXSetting";
   private const string MousePositionYSetting = "MousePositionYSetting";
   private const string RevivedState = "0";
   private const string HiddenState = "1";

   private static bool ForceReviveRequestedCache = false;
   private static bool FocusModeRequestedCache = false;
   private static bool SweepModeRequestedCache = false;
   private static IntPtr ActiveWindowCache = IntPtr.Zero;

   private static readonly uint RESOLUTION_4K_WIDTH = 3840;
   private static readonly uint RESOLUTION_4K_HEIGHT = 2160;

   private static readonly uint RESOLUTION_2K_WIDTH = 2560;
   private static readonly uint RESOLUTION_2K_HEIGHT = 1440;
   private static readonly uint MOUSE_RESTORE_THRESHOLD = 200;
   private static readonly int SWEEP_SNAP_THRESHOLD = 150;
   private static readonly double SWEEP_NO_RESIZE_THRESHOLD = 0.9;
   private static readonly int TASKBAR_HEIGHT = 40;

   public static string KEY_SHIFT = "16";
   public static string KEY_CTRL = "17";
   public static string KEY_ALT = "18";

   private static readonly bool enableMouseMove = true;
   private static readonly bool enableDebugPrints = true;
   private static readonly bool prioritizeHidingDefault = true;
   private static readonly bool keepRestoringDefault = true;
   private static readonly bool enableForceRevive = true;
   private static readonly bool enableFocusMode = true;
   private static readonly bool enableSweepMode = true;
   private static readonly bool enableSweepModeSnap = true;
   private static readonly bool debugPrintHideRevive = enableDebugPrints && false;
   private static readonly bool debugPrintStartStop = enableDebugPrints && false;
   private static readonly bool debugPrintFindMonitorId = enableDebugPrints && false;
   private static readonly bool debugPrintNoMonitorFound = enableDebugPrints && true;
   private static readonly bool debugWindowFiltering = enableDebugPrints && false;
   private static readonly bool debugPrintMoveCursor = enableDebugPrints && false;
   private static readonly bool debugPrintDecideMinRevive = enableDebugPrints && false;
   private static readonly bool debugPrintCountToMin = enableDebugPrints && false;
   private static readonly bool debugPrintForceReviveKey = enableDebugPrints && false;
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
      CacheAll();
      bool windowsAreVisible = CountWindowsToHide() > 0;
      bool windowsWereMinimized = WereWindowsMinimized();
      bool windowsWereSwept = WereWindowsSwept();
      bool windowsWereHiddenEver = windowsWereMinimized || windowsWereSwept;
      bool windowsWereHiddenCurrentMode = (!IsSweepModeRequested() && windowsWereMinimized) ||
                                          (IsSweepModeRequested() && windowsWereSwept);

      string debugInfo = $"IsForceReviveRequested: {IsForceReviveRequested()}\n" +
                         $"IsFocusModeRequested: {IsFocusModeRequested()}\n" +
                         $"IsSweepModeRequested: {IsSweepModeRequested()}\n\n" +
                         $"windowsAreVisible: {windowsAreVisible}\n" +
                         $"windowsWereMinimized: {windowsWereMinimized}\n" +
                         $"windowsWereSwept: {windowsWereSwept}\n\n" +
                         $"windowsWereHiddenEver: {windowsWereHiddenEver}\n" +
                         $"windowsWereHiddenCurrentMode: {windowsWereHiddenCurrentMode}";

      if (IsForceReviveRequested())
      {
         if (debugPrintDecideMinRevive) MessageBox.Show($"Force revive\n\n{debugInfo}");
         // bool forceAllWindowsRevive = !windowsWereHidden;
         HandleReviving(false);
         // todo ShouldKeepRevivingOnForce()??? inside HandleReviving?
      }
      else // no forceRevive
      {
         if (windowsAreVisible && !windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"Visible and NOT hidden\n\n{debugInfo}");
            HandleHiding();
         }
         else if (windowsAreVisible && windowsWereHiddenCurrentMode)
         {
            if (ShouldPrioritizeHiding())
            {
               if (debugPrintDecideMinRevive) MessageBox.Show($"Visible and hidden (prio hide)\n\n{debugInfo}");
               HandleHiding(); // ignore saved windows and hide visible ones
            }
            else
            {
               if (debugPrintDecideMinRevive) MessageBox.Show($"Visible and hidden (NO prio hide)\n\n{debugInfo}");
               HandleReviving(false); // ignore visible windows and revive saved ones
               // todo is ShouldKeepRevivingOnForce() good idea here?
            }
         }
         else if (!windowsAreVisible && windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"NOT visible and hidden\n\n{debugInfo}");
            HandleReviving(false); // revive saved windows only
         }
         else if (!windowsAreVisible && !windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"NOT visible and NOT hidden\n\nDoing Nothing!\n\n{debugInfo}");
            // HandleReviving(true); // force revive all windows
            // do nothing 
         }
         else
         {
            MessageBox.Show($"ERROR else state not expected!");
         }
      }


      // if (forceRevive) // modifier key is pressed, do not hide windows
      // {
      //    if (debugPrintDecideMinRevive) MessageBox.Show($"force revive, windows were" + (windowsWereMinimized ? "" : "NOT ") + " minimized previously");
      //    // try reviving only saved windows first if there were any
      //    int revivedCount = HandleReviving(!windowsWereMinimized);
      //    // optionally restore all other minimized if saved windows were manually restored
      //    if ((revivedCount < 1) && ShouldKeepRevivingOnForce() && windowsWereMinimized)
      //    {
      //       revivedCount = HandleReviving(true); // restore all windows (not only saved but minimized manually)
      //       if (revivedCount < 1) MessageBox.Show($"There is no other windows to restore (todo remove)");
      //    }
      // }
      // else if (!windowsWereMinimized && windowsAreVisible)
      // {
      //    // nothing was previously minimized but there are visible windows, just minimize them
      //    if (debugPrintDecideMinRevive) MessageBox.Show($"NOT windowsWereMinimized && windowsAreVisible");
      //    int hiddenCount = HandleHiding();
      //    if (hiddenCount < 1) MessageBox.Show($"ERROR no windows were minimized but should be!");
      // }
      // else if (!windowsWereMinimized && !windowsAreVisible)
      // {
      //    // nothing was previously minimized, nothing is visible, force all windows that may have been manually minimized to restore
      //    if (debugPrintDecideMinRevive) MessageBox.Show($"NOT windowsWereMinimized && NOT windowsAreVisible");
      //    int revivedCount = HandleReviving(true); // force revive all windows
      //    if (revivedCount < 1) MessageBox.Show($"No windows were restored but that may be ok if there was none at all (todo remove)"); // todo remove
      // }
      // else if (windowsWereMinimized && windowsAreVisible)
      // {
      //    // there were windows minimized but there are also manually restored or new windows visible
      //    // decide if we should restore saved windows or minimize visible ones
      //    if (ShouldPrioritizeHiding())
      //    {
      //       if (debugPrintDecideMinRevive) MessageBox.Show($"windowsWereMinimized && windowsAreVisible && ShouldPrioritizeHiding");
      //       int hiddenCount = HandleHiding();
      //       if (hiddenCount < 1) MessageBox.Show($"ERROR no windows were minimized but should be (prio min)!");
      //    }
      //    else // don't prioritize minimizing new windows and restore saved ones
      //    {
      //       if (debugPrintDecideMinRevive) MessageBox.Show($"windowsWereMinimized && windowsAreVisible && NOT ShouldPrioritizeHiding");
      //       int revivedCount = HandleReviving(false); // try reviving only saved windows
      //       if ((revivedCount < 1) && ShouldKeepRevivingOnForce()) // optionally restore all other minimized if saved windows were manually restored
      //       {
      //          revivedCount = HandleReviving(true); // restore all windows (not only saved but minimalized manually)
      //          if (revivedCount < 1) MessageBox.Show($"There is no other windows to restore (todo remove)");
      //       }
      //    }
      // }
      // else if (windowsWereMinimized && !windowsAreVisible)
      // {
      //    // there were windows minimized and there are no new visible windows, just restore saved windows
      //    if (debugPrintDecideMinRevive) MessageBox.Show($"windowsWereMinimized && NOT windowsAreVisible");
      //    int revivedCount = HandleReviving(false); // restore only saved windows
      //    if (revivedCount < 1) MessageBox.Show($"ERROR no windows were restored but should be");
      // }
      // else
      // {
      //    MessageBox.Show($"ERROR else state not expected!");
      // }
   }

   private static bool WereWindowsMinimized()
   {
      string setting = BFS.ScriptSettings.ReadValue(ScriptStateMinSetting);
      return !string.IsNullOrEmpty(setting) && (setting.Equals(HiddenState, StringComparison.Ordinal));
   }

   private static bool WereWindowsSwept()
   {
      string setting = BFS.ScriptSettings.ReadValue(ScriptStateSweepSetting);
      return !string.IsNullOrEmpty(setting) && (setting.Equals(HiddenState, StringComparison.Ordinal));
   }

   public static void HandleHiding()
   {
      if (debugPrintStartStop) MessageBox.Show("start HIDE");

      // get monitor ID and bounds of OLED monitor
      uint monitorIdOled = GetOledMonitorID();
      // Rectangle monitorBoundsOled = BFS.Monitor.GetMonitorBoundsByID(monitorIdOled);

      // get windows to be hidden
      IntPtr[] windowsToHide = GetFilteredVisibleWindows(monitorIdOled);

      // save handle to currently active window
      IntPtr activeWindowHandle = GetCachedActiveWindow();
      if (debugPrintFocusMode) MessageBox.Show($"focus window found: {BFS.Window.GetText(activeWindowHandle)}");

      int count = 0;
      if (IsSweepModeRequested())
      {
         count = SweepWindows(windowsToHide, activeWindowHandle);
      }
      else
      {
         count = MinimizeWindows(windowsToHide, activeWindowHandle);
      }

      // hide mouse cursor to primary monitor
      if (ShoudlMoveMouse() && (count > 0)) HandleMouseOut();
   }

   public static int MinimizeWindows(IntPtr[] windowsToMinimize, IntPtr activeWindowHandle)
   {
      // this will store the windows that we are minimizing so we can restore them later
      string minimizedWindowsSaveList = "";

      int minimizedWindowsCount = 0;
      // loop through all the visible windows on the monitor
      foreach (IntPtr window in windowsToMinimize)
      {
         // if focus mode enabled skip focused window from list to minimize
         if (IsFocusModeRequested() && window == activeWindowHandle)
         {
            if (debugPrintHideRevive) MessageBox.Show($"skipping focused window {BFS.Window.GetText(window)}");
            minimizedWindowsCount += 1; // treat active window in focus mode as if it was minimized
            continue;
         }

         if (debugPrintHideRevive) MessageBox.Show($"minimizing window {BFS.Window.GetText(window)}");
         WindowUtils.MinimizeWindow(window);

         // use variables for both minimizing and sweeping
         minimizedWindowsCount += 1;
         // add the window to the list of minimized windows
         minimizedWindowsSaveList += window.ToInt64().ToString() + "|";
      }

      // change focus to desktop only when not in focus mode
      if (!IsFocusModeRequested())
      {
         // it is a fix in order to enable alt-tabbing back to top minimized window
         // and being able to restore minimized windows from taskbar with only 1 mouse click
         WindowUtils.FocusOnDekstop();
      }

      // save the list of windows that were minimized
      BFS.ScriptSettings.WriteValue(MinimizedWindowsListSetting, minimizedWindowsSaveList);

      // set the script state to HiddenState
      BFS.ScriptSettings.WriteValue(ScriptStateMinSetting, HiddenState);

      if (debugPrintStartStop) MessageBox.Show($"finished MIN (minimized {minimizedWindowsCount}/{windowsToMinimize.Length} windows)");

      return minimizedWindowsCount;
   }

   public static int SweepWindows(IntPtr[] windowsToSweep, IntPtr activeWindowHandle)
   {
      // get bounds of sweep-target monitor
      Rectangle monitorBoundsSweep = BFS.Monitor.GetMonitorBoundsByID(GetSweepTargetMonitorID());

      // sweep mode in reverse order (the same as restoring) compared to minimizing windows
      Array.Reverse(windowsToSweep);

      // this will store the windows that we are sweepping so we can revive them later
      string sweptWindowsSaveList = "";

      int sweptWindowsCount = 0;
      // loop through all the visible windows on the monitor
      foreach (IntPtr windowHandle in windowsToSweep)
      {
         // if focus mode enabled skip focused windowHandle from list to minimize
         if (IsFocusModeRequested() && windowHandle == activeWindowHandle)
         {
            if (debugPrintHideRevive) MessageBox.Show($"skipping focused window {BFS.Window.GetText(windowHandle)}");
            sweptWindowsCount += 1; // treat active window in focus mode as if it was swept
            continue;
         }

         if (debugPrintHideRevive) MessageBox.Show($"sweeping window {BFS.Window.GetText(windowHandle)}");
         SweepOneWindow(windowHandle, /*monitorBoundsOled,*/ monitorBoundsSweep);

         sweptWindowsCount += 1;
         // add the window to the list of swept windows
         sweptWindowsSaveList += windowHandle.ToInt64().ToString() + "|";
      }

      if (IsFocusModeRequested())
      {
         // restore focus from swept windows (which got focus because of pushing on top for Z-reordering)
         WindowUtils.FocusOnWindow(activeWindowHandle);
      }
      // focus and move mouse only when not in focus mode and not sweeping
      else /*if (!sweepMode)*/
      {
         // it is a fix in order to enable alt-tabbing back to top minimized window
         // and being able to restore minimized windows from taskbar with only 1 mouse click
         // WindowUtils.FocusOnDekstop();
      }

      // save the list of windows that were swept
      BFS.ScriptSettings.WriteValue(SweptdWindowsListSetting, sweptWindowsSaveList);

      // set the script sweep state to HiddenState
      BFS.ScriptSettings.WriteValue(ScriptStateSweepSetting, HiddenState);

      if (debugPrintStartStop) MessageBox.Show($"finished SWEEP (swept {sweptWindowsCount}/{windowsToSweep.Length} windows)");

      return sweptWindowsCount;
   }

   public static void HandleReviving(bool reviveAll)
   {
      if (debugPrintStartStop) MessageBox.Show("start REVIVE");

      // first restore mouse cursor position if enabled
      if (ShoudlMoveMouse()) HandleMouseBack();

      // get windows to be revived
      IntPtr[] windowsToRevive = GetListOfWindowsToRevive(reviveAll);

      // save list of currently visible windows to later push them on top of revived ones
      IntPtr[] windowsToPushOnTop = new IntPtr[] { };
      if (IsFocusModeRequested())
      {
         windowsToPushOnTop = GetFilteredVisibleWindows(GetOledMonitorID());
      }

      // select proper hiding function based on requested mode
      Func<IntPtr[], int> revivingFunction = IsSweepModeRequested() ? UnsweepWindows : RestoreWindows;

      // revive windows with correct function
      int count = revivingFunction(windowsToRevive);

      // try restoring all windows if restoring saved ones failed (if enabled)
      bool keepReviving = !reviveAll && count < 1 && ShouldKeepRevivingOnForce();
      if (keepReviving)
      {
         windowsToRevive = GetListOfWindowsToRevive(true);
         count = revivingFunction(windowsToRevive);
      }

      // if in focus mode, push windows that were active on top of revived ones
      FixZorderAfterReviveInFocusMode(windowsToPushOnTop);

      if (debugPrintStartStop) MessageBox.Show($"finished REVIVE (revived {count}/{windowsToRevive.Length} windows)\n\n" +
                                               $"keptReviving: {keepReviving}");

   }

   public static IntPtr[] GetListOfWindowsToRevive(bool forceReviveAll)
   {
      if (IsSweepModeRequested())
      {
         return GetListOfWindowsToUnsweep(forceReviveAll);
      }
      else
      {
         return GetListOfWindowsToRestore(forceReviveAll);
      }
   }

   public static IntPtr[] GetListOfWindowsToRestore(bool forceReviveAll)
   {
      List<IntPtr> windowsToRestore = new List<IntPtr>();

      if (forceReviveAll) // restore all minimized windows on OLED monitor
      {
         windowsToRestore = GetFilteredMinimizedWindows(GetOledMonitorID()).ToList();
      }
      else // only restore windows previously minimized
      {
         // get the windows that we minimized previously
         string savedWindows = BFS.ScriptSettings.ReadValue(MinimizedWindowsListSetting);

         string[] windowsToReviveStrings = savedWindows.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

         // restore windows in reverse order than minimizing
         Array.Reverse(windowsToReviveStrings);

         // parse window handles from string to int
         foreach (string window in windowsToReviveStrings)
         {
            // try to turn the string into a long value if we can't convert it, go to the next setting
            long windowHandleValue;
            if (!Int64.TryParse(window, out windowHandleValue))
               continue;

            windowsToRestore.Add(new IntPtr(windowHandleValue));
         }
      }

      return windowsToRestore.ToArray();
   }

   public static IntPtr[] GetListOfWindowsToUnsweep(bool forceReviveAll)
   {
      List<IntPtr> windowsToUnsweep = new List<IntPtr>();

      if (forceReviveAll) // "unsweep" all windows on sweep-target monitor 
      {
         // forceRevive for sweep is moving all windows from sweep-target to OLED regardless whether they were ever swept or not
         windowsToUnsweep = GetFilteredVisibleWindows(GetSweepTargetMonitorID()).ToList();
      }
      else // only unsweep windows previously swept
      {
         // get the windows that we swept previously
         string savedWindows = BFS.ScriptSettings.ReadValue(SweptdWindowsListSetting);

         string[] windowsToReviveStrings = savedWindows.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
         //MessageBox.Show($"GetListOfWindowsToUnsweep windowsToReviveStrings:\n{string.Join(", ", windowsToReviveStrings)}"); // todo

         // // unsweep windows in reverse order than minimizing
         // Array.Reverse(windowsToReviveStrings);

         // parse window handles from string to int
         // todo add saving and parsing original window position
         foreach (string window in windowsToReviveStrings)
         {
            // try to turn the string into a long value if we can't convert it, go to the next setting
            long windowHandleValue;
            if (!Int64.TryParse(window, out windowHandleValue)) continue;

            // check if saved window handle is still valid (e.g. window wasn't closed in the meantime)
            if (!WindowUtils.IsWindowValid(new IntPtr(windowHandleValue))) continue;

            //MessageBox.Show($"GetListOfWindowsToUnsweep\nstring: {string.Join(", ", windowsToReviveStrings)}\n\nadding:\n{windowHandleValue}\nparsed from: {window}"); // todo
            windowsToUnsweep.Add(new IntPtr(windowHandleValue));
         }
      }

      return windowsToUnsweep.ToArray();
   }

   public static int RestoreWindows(IntPtr[] windowsToRestore)
   {
      int restoredWindowsCount = 0;
      // loop through each window to restore
      foreach (IntPtr windowHandle in windowsToRestore)
      {
         if (BFS.Window.IsMinimized(windowHandle))
         {
            if (debugPrintHideRevive) MessageBox.Show($"restoring window {BFS.Window.GetText(new IntPtr(windowHandle))}");
            WindowUtils.RestoreWindow(windowHandle);
            WindowUtils.PushToTop(windowHandle);
            restoredWindowsCount += 1;
         }
         else
         {
            if (debugPrintHideRevive) MessageBox.Show($"already restored window {BFS.Window.GetText(new IntPtr(windowHandle))}");
         }
      }

      // clear the windows that we saved
      BFS.ScriptSettings.WriteValue(MinimizedWindowsListSetting, string.Empty);

      // set the script to RevivedState
      BFS.ScriptSettings.WriteValue(ScriptStateMinSetting, RevivedState);
      return restoredWindowsCount;

   }

   public static int UnsweepWindows(IntPtr[] windowsToUnsweep)
   {
      // todo fix unsweeping saved but closed (not existing) windows
      // get bounds of OLED (un-sweep-targer) monitor
      Rectangle monitorBoundsUnsweep = BFS.Monitor.GetMonitorBoundsByID(GetOledMonitorID());

      int unsweptWindowsCount = 0;
      // loop through each window to restore
      foreach (IntPtr windowHandle in windowsToUnsweep)
      {
         if (debugPrintHideRevive) MessageBox.Show($"[TODO] unsweeping window {BFS.Window.GetText(new IntPtr(windowHandle))}");

         // todo how to unsweep??
         SweepOneWindow(windowHandle, monitorBoundsUnsweep);
         unsweptWindowsCount += 1;
      }

      // clear the windows that we saved
      BFS.ScriptSettings.WriteValue(SweptdWindowsListSetting, string.Empty);

      // set the script to RevivedState
      BFS.ScriptSettings.WriteValue(ScriptStateSweepSetting, RevivedState);
      return unsweptWindowsCount;

   }

   public static void FixZorderAfterReviveInFocusMode(IntPtr[] windowsToPushOnTop)
   {
      if (IsFocusModeRequested() && windowsToPushOnTop.Length > 0)
      {
         Array.Reverse(windowsToPushOnTop);
         foreach (IntPtr windowHandle in windowsToPushOnTop)
         {
            if (debugPrintHideRevive) MessageBox.Show($"Pushing window {BFS.Window.GetText(new IntPtr(windowHandle))} on top (focus mode)");
            WindowUtils.PushToTop(windowHandle);
         }
      }
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

   public static void SweepOneWindow(IntPtr windowHandle,/* Rectangle boundsFrom,*/ Rectangle boundsTo)
   {
      Rectangle boundsFrom = BFS.Monitor.GetMonitorBoundsByWindow(windowHandle);
      if (debugPrintSweepMode) MessageBox.Show($"SweepOneWindow:\nboundsFrom: {boundsFrom},\nboundsTo {boundsTo}");

      // when window is already maximized restore it (so that it can be maximized on target screen)
      // and treat it as if it's size would be the same as source monitor 
      // (because GetBounds() function gives you window bounds from before maximizing it)
      Rectangle windowBounds = new Rectangle() { };
      if (BFS.Window.IsMaximized(windowHandle))
      {
         BFS.Window.Restore(windowHandle);
         windowBounds = boundsFrom;
      }
      else
      {
         windowBounds = WindowUtils.GetBounds(windowHandle);
      }

      var result = CalculateSweptWindowPos(windowBounds, boundsFrom, boundsTo);
      Rectangle newPos = result.newWindowBounds;
      WindowUtils.SetSizeAndLocation(windowHandle, newPos.X, newPos.Y, newPos.Width, newPos.Height);
      WindowUtils.PushToTop(windowHandle);

      // maximize window when it is big enough
      if (result.shouldMaximize)
      {
         System.Threading.Thread.Sleep(20);
         WindowUtils.MaximizeWindow(windowHandle);
      }
   }

   public static (bool shouldMaximize, Rectangle newWindowBounds) CalculateSweptWindowPos(Rectangle boundsWindow, Rectangle boundsFrom, Rectangle boundsTo)
   {
      Rectangle newWindowBounds = new Rectangle();
      bool shouldMaximize = false;

      int maxWidth = (int)(SWEEP_NO_RESIZE_THRESHOLD * boundsTo.Width);
      int maxHeight = (int)(SWEEP_NO_RESIZE_THRESHOLD * (boundsTo.Height - TASKBAR_HEIGHT));

      if ((boundsWindow.Width < maxWidth) ||
          (boundsWindow.Height < maxHeight))
      {
         // small window in at least one direction, don't change size for now
         newWindowBounds.Width = boundsWindow.Width;
         newWindowBounds.Height = boundsWindow.Height;

         // calculate how far is the windows from old monitor borders (ignore taskbar size because of autohide)
         int leftDistOld = boundsWindow.X - boundsFrom.X + 1;
         int rightDistOld = (boundsFrom.X + boundsFrom.Width) - (boundsWindow.X + boundsWindow.Width) + 1;
         int topDistOld = boundsWindow.Y - boundsFrom.Y + 1;
         int bottomDistOld = (boundsFrom.Y + boundsFrom.Height) - (boundsWindow.Y + boundsWindow.Height) + 1;

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
         // set size to max bounds below which windows are not resized
         // this will set the size to which window will restore when unmaximizing
         newWindowBounds.X = boundsTo.X + (boundsTo.Width - maxWidth) / 2;
         newWindowBounds.Y = boundsTo.Y + (boundsTo.Height - maxHeight) / 2;
         newWindowBounds.Width = maxWidth;
         newWindowBounds.Height = maxHeight;
         shouldMaximize = true;

         if (debugPrintSweepModeCalcPos) MessageBox.Show($"\nboundsWindow\t{boundsWindow}\n" +
                                             $"boundsFrom\t{boundsFrom}\n" +
                                             $"boundsTo\t{boundsTo}\n\n" +
                                             $"newWindowBounds\t{newWindowBounds}\n" +
                                             $"maxWidth\t{maxWidth}\tmaxHeight\t{maxHeight}\n" +
                                             $"shouldMaximize:\t{shouldMaximize}");
      }

      return (shouldMaximize, newWindowBounds);
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
            if (debugPrintFindMonitorId) MessageBox.Show($"found 2k monitor with ID: {id}\nbounds: {bounds}");
            return id;
         }
      }

      MessageBox.Show($"ERROR! did not find monitor with 2K resolution");
      return UInt32.MaxValue;
   }

   public static int CountWindowsToHide()
   {
      IntPtr[] windowsToHide = GetFilteredVisibleWindows(GetOledMonitorID());

      // save handle to currently active window
      IntPtr activeWindowHandle = GetCachedActiveWindow();
      int activeWindowOffset = 0;

      if (debugPrintCountToMin) MessageBox.Show($"CountWindowsToHide: {windowsToHide.Length}");
      foreach (IntPtr window in windowsToHide)
      {
         if (IsFocusModeRequested() && window == activeWindowHandle)
         {
            if (debugPrintCountToMin) MessageBox.Show($"Skipping counting focused window as window to hide: \n\n" +
                                                      $"|{BFS.Window.GetText(window)}|\n\n|{BFS.Window.GetClass(window)}|");
            activeWindowOffset = -1;
         }
         else
         {
            if (debugPrintCountToMin) MessageBox.Show($"To hide: \n\n|{BFS.Window.GetText(window)}|\n\n|{BFS.Window.GetClass(window)}|");

         }
      }

      return windowsToHide.Length + activeWindowOffset;
   }

   public static bool ShouldPrioritizeHiding()
   {
      // todo store in settings
      return prioritizeHidingDefault;
   }

   public static bool ShouldKeepRevivingOnForce()
   {
      // todo store in settings
      return keepRestoringDefault && IsForceReviveRequested();
   }

   public static bool ShoudlMoveMouse()
   {
      return enableMouseMove && !IsFocusModeRequested();
   }

   public static void CacheAll()
   {
      CacheForceReviveRequeste();
      CacheFocusModeRequest();
      CacheSweepModeRequest();
      CacheActiveWindow();
   }
   public static void CacheForceReviveRequeste()
   {
      // bool ForceReviveRequestedCache = BFS.Input.IsMouseDown("1;");
      ForceReviveRequestedCache = BFS.Input.IsKeyDown(KEY_SHIFT);
      if (debugPrintForceReviveKey) MessageBox.Show($"ForceRevive key is" + (ForceReviveRequestedCache ? "" : " NOT") + " pressed, caching");
   }

   public static void CacheFocusModeRequest()
   {
      // bool FocusModeRequestedCache = BFS.Input.IsMouseDown("2;");
      FocusModeRequestedCache = BFS.Input.IsKeyDown(KEY_CTRL);
      if (debugPrintFocusModeKey) MessageBox.Show($"FocusMode key is" + (FocusModeRequestedCache ? "" : " NOT") + " pressed, caching");
   }

   public static void CacheSweepModeRequest()
   {
      // bool SweepModeRequestedCache = BFS.Input.IsMouseDown("2;");
      SweepModeRequestedCache = BFS.Input.IsKeyDown(KEY_ALT);
      if (debugPrintSweepModeKey) MessageBox.Show($"SweepMode key is" + (SweepModeRequestedCache ? "" : " NOT") + " pressed, caching");
   }

   public static void CacheActiveWindow()
   {
      ActiveWindowCache = BFS.Window.GetFocusedWindow();
   }
   public static bool IsForceReviveRequested()
   {
      if (debugPrintForceReviveKey) MessageBox.Show($"ForceRevive cache: {ForceReviveRequestedCache}");
      return ForceReviveRequestedCache;
   }

   public static bool IsFocusModeRequested()
   {
      if (debugPrintFocusModeKey) MessageBox.Show($"FocusMode cache: {FocusModeRequestedCache}");
      return FocusModeRequestedCache;
   }

   public static bool IsSweepModeRequested()
   {
      if (debugPrintSweepModeKey) MessageBox.Show($"SweepMode cache: {SweepModeRequestedCache}");
      return SweepModeRequestedCache;
   }

   public static IntPtr GetCachedActiveWindow()
   {
      return ActiveWindowCache;
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

      [DllImport("user32.dll", SetLastError = true)]
      [return: MarshalAs(UnmanagedType.Bool)]
      public static extern bool IsWindow(IntPtr hWnd);


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
                        $"windowHandle: |{windowHandle}|\n\n" +
                        $"text: |{text}|\n\n" +
                        $"requested pos: x.{x} y.{y} w.{w} h.{h}");
         }

         if (!GetWindowRect(windowHandle, out includeShadow)) // including shadow
         {
            int errorCode = Marshal.GetLastWin32Error();
            string text = BFS.Window.GetText(windowHandle);
            MessageBox.Show($"ERROR CompensateForShadow-GetWindowRect windows API: {errorCode}\n\n" +
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
         FocusOnWindow(hWndDesktop);
      }

      public static void FocusOnWindow(IntPtr windowHandle)
      {
         bool success = SetForegroundWindow(windowHandle);
         if (!success)
         {
            int errorCode = Marshal.GetLastWin32Error();
            MessageBox.Show($"FocusOnWindow error: {errorCode}");
         }
      }

      public static void PushToTop(IntPtr windowHandle)
      {
         if (debugPrintHideRevive) MessageBox.Show($"pushing\n|{BFS.Window.GetText(windowHandle)}|\non top");
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
}