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
using System.Collections.Specialized;
public static class DisplayFusionFunction
{
   private const string MinimizedWindowsListSetting = "OledMinimizerMinimizedWindowsList";
   private const string SweptWindowsListSetting = "OledMinimizerSweptWindowsList";
   private const string MousePositionXSetting = "MousePositionXSetting";
   private const string MousePositionYSetting = "MousePositionYSetting";
   private const uint RESOLUTION_4K_WIDTH = 3840;
   private const uint RESOLUTION_4K_HEIGHT = 2160;

   private const uint RESOLUTION_2K_WIDTH = 2560;
   private const uint RESOLUTION_2K_HEIGHT = 1440;
   private const uint MOUSE_RESTORE_THRESHOLD = 200;
   private const int SWEEP_SNAP_THRESHOLD = 150;
   private const int SWEEP_SNAP_HALFSPLIT_THRESHOLD = 500;
   private const double SWEEP_NO_RESIZE_THRESHOLD = 0.9;
   private const int TASKBAR_HEIGHT = 40;

   public const string KEY_SHIFT = "16";
   public const string KEY_CTRL = "17";
   public const string KEY_ALT = "18";

   private const bool enableMouseMove = true;
   private const bool enableDebugPrints = true;
   private const bool prioritizeHidingDefault = true;
   private const bool keepRestoringDefault = true;
   private const bool enableForceRevive = true;
   private const bool enableFocusMode = true;
   private const bool enableSweepMode = true;
   private const bool enableSweepModeSnap = true;
   private const bool enableSweepModeSnapHalfSplit = true;
   private const bool enableBoundingBoxMode = true;
   private const bool debugPrintHideRevive = enableDebugPrints && false;
   private const bool debugPrintStartStop = enableDebugPrints && false;
   private const bool debugPrintFindMonitorId = enableDebugPrints && false;
   private const bool debugPrintNoMonitorFound = enableDebugPrints && true;
   private const bool debugWindowFiltering = enableDebugPrints && false;
   private const bool debugPrintMoveCursor = enableDebugPrints && false;
   private const bool debugPrintDecideMinRevive = enableDebugPrints && false;
   private const bool debugPrintCountToMin = enableDebugPrints && false;
   private const bool debugPrintForceReviveKey = enableDebugPrints && false;
   private const bool debugPrintFocusModeKey = enableDebugPrints && false;
   private const bool debugPrintFocusMode = enableDebugPrints && false;
   private const bool debugPrintSweepModeKey = enableDebugPrints && false;
   private const bool debugPrintSweepMode = enableDebugPrints && false;
   private const bool debugPrintSweepModeCalcPos = enableDebugPrints && false;
   private const bool debugPrintListOfWindows = enableDebugPrints && false;

   private static string listOfWindowsToUnsweepStr = ""; // debug only
   private static string listOfWindowsToHideStr = ""; // debug only

   private static bool ForceReviveRequestedCache = false;
   private static bool FocusModeRequestedCache = false;
   private static bool SweepModeRequestedCache = false;
   private static IntPtr ActiveWindowHandleCache = IntPtr.Zero;

   private static OrderedDictionary unsweepWindowsInfoMap = new();


   private static List<string> classnameBlacklist = new List<string> {"DFTaskbar", "DFTitleBarWindow", "Shell_TrayWnd",
                                                                       "tooltips", "Shell_InputSwitchTopLevelWindow",
                                                                       "Windows.UI.Core.CoreWindow", "Progman", "SizeTipClass",
                                                                       "DF", "WorkerW", "SearchPane", "KbxLabelClass",
                                                                       "WindowsForms10.tooltips_class", "Xaml_WindowedPopupClass",
                                                                       "Ghost"};
   private static List<string> textBlacklist = new List<string> {"Program Manager", "Volume Mixer", "Snap Assist", "Greenshot capture form",
                                                                  "Battery Information", "Date and Time Information", "Network Connections",
                                                                  "Volume Control", "Start", "Search", "SubFolderTipWindow"};
   private static StringBuilder saveInfoStrBuilder = new();

   public static void Run(IntPtr windowHandle)
   {
      Init();

      IntPtr[] windowsToHide = GetListOfWindowsToHide(GetOledMonitorID());

      bool windowsToHidePresent = windowsToHide.Length > 0;
      bool windowsWereMinimized = WereWindowsMinimized();
      bool windowsWereSwept = WereWindowsSwept();
      bool windowsWereHiddenCurrentMode = (!IsSweepModeRequested() && windowsWereMinimized) ||
                                          (IsSweepModeRequested() && windowsWereSwept);

      string debugInfo = $"IsForceReviveRequested: {IsForceReviveRequested()}\n" +
                         $"IsFocusModeRequested: {IsFocusModeRequested()}\n" +
                         $"IsSweepModeRequested: {IsSweepModeRequested()}\n\n" +
                         $"windowsToHidePresent: {windowsToHidePresent}\n" +
                         $"windowsWereMinimized: {windowsWereMinimized}\n" +
                         $"windowsWereSwept: {windowsWereSwept}\n\n" +
                         $"windowsWereHiddenCurrentMode: {windowsWereHiddenCurrentMode}";

      if (IsForceReviveRequested())
      {
         if (debugPrintDecideMinRevive) MessageBox.Show($"Force revive\n\n{debugInfo}");
         HandleReviving(false);
      }
      else // no forceRevive
      {
         if (windowsToHidePresent && !windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"NOTHING to hide and NOT hidden\n\n{debugInfo}");
            HandleHiding(windowsToHide);
         }
         else if (windowsToHidePresent && windowsWereHiddenCurrentMode)
         {
            if (ShouldPrioritizeHiding())
            {
               if (debugPrintDecideMinRevive) MessageBox.Show($"PRESENT to hide and WERE hidden (prio hide)\n\n{debugInfo}");
               HandleHiding(windowsToHide); // ignore saved windows and hide visible ones
            }
            else
            {
               if (debugPrintDecideMinRevive) MessageBox.Show($"PRESENT to hide and WERE hidden (NO prio hide)\n\n{debugInfo}");
               HandleReviving(false); // ignore visible windows and revive saved ones
            }
         }
         else if (!windowsToHidePresent && windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"NOTHING to hide and WERE hidden\n\n{debugInfo}");
            HandleReviving(false); // revive saved windows only
         }
         else if (!windowsToHidePresent && !windowsWereHiddenCurrentMode)
         {
            if (debugPrintDecideMinRevive) MessageBox.Show($"NOTHING to hide and NOT hidden\n\nDoing Nothing!\n\n{debugInfo}");
            // HandleReviving(true); // todo force revive all windows or do nothing?
         }
         else
         {
            MessageBox.Show($"ERROR <min> else state not expected!");
         }
      }

      if (debugPrintListOfWindows) MessageBox.Show($"END\n\nlistofwindowsto unsweep:\n\n{listOfWindowsToUnsweepStr}\n\n" +
                                                  $"listofwindowsto hide:\n\n{listOfWindowsToHideStr}");
   }

   private static bool WereWindowsMinimized()
   {
      string setting = BFS.ScriptSettings.ReadValue(MinimizedWindowsListSetting);
      return !string.IsNullOrEmpty(setting);
   }

   private static bool WereWindowsSwept()
   {
      string setting = BFS.ScriptSettings.ReadValue(SweptWindowsListSetting);
      return !string.IsNullOrEmpty(setting);
   }

   public static void HandleHiding(IntPtr[] windowsToHide)
   {
      if (debugPrintStartStop) MessageBox.Show("start HIDE");

      int count = 0;
      if (IsSweepModeRequested())
      {
         count = SweepWindows(windowsToHide);
      }
      else
      {
         count = MinimizeWindows(windowsToHide);
      }

      // hide mouse cursor to primary monitor
      if (ShoudlMoveMouse() && (count > 0)) HandleMouseOut();
   }

   public static int MinimizeWindows(IntPtr[] windowsToMinimize)
   {
      // this will store the windows that we are minimizing so we can restore them later
      string minimizedWindowsSaveList = "";

      // loop through all windows to minimize
      int minimizedWindowsCount = 0;
      foreach (IntPtr windowHandle in windowsToMinimize)
      {
         if (DetectedInvalidWindow(windowHandle, "MinimizeWindows", $"minimizedWindowsCount: {minimizedWindowsCount}")) continue;
         if (debugPrintHideRevive) MessageBox.Show($"minimizing window {BFS.Window.GetText(windowHandle)}");

         WindowUtils.MinimizeWindow(windowHandle);
         minimizedWindowsCount += 1;

         // add the window to the list of minimized windows
         minimizedWindowsSaveList += windowHandle.ToInt64().ToString() + "|";
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

      if (debugPrintStartStop) MessageBox.Show($"finished MIN (minimized {minimizedWindowsCount}/{windowsToMinimize.Length} windows)");

      return minimizedWindowsCount;
   }

   public static int SweepWindows(IntPtr[] windowsToSweep)
   {
      if (windowsToSweep.Length == 0)
      {
         MessageBox.Show($"ERROR <min>! SweepWindows no windows on sweep list");
         return 0;
      }

      // get bounds of monitors
      Rectangle monitorBoundsTarget = BFS.Monitor.GetMonitorBoundsByID(GetSweepTargetMonitorID());
      Rectangle monitorBoundsSource = BFS.Monitor.GetMonitorBoundsByID(GetOledMonitorID());

      // sweep mode in reverse order (the same as restoring) compared to minimizing windows
      Array.Reverse(windowsToSweep);

      // calculate bounding box of all windows
      Rectangle boundingBox = CalculateBoundingBox(windowsToSweep);

      // loop through all windows to be swept
      int sweptWindowsCount = 0;
      foreach (IntPtr windowHandle in windowsToSweep)
      {
         if (debugPrintHideRevive) MessageBox.Show($"sweeping window {BFS.Window.GetText(windowHandle)}");

         // save info about swept window to store in settings so we can unsweep them later
         AppendSavedWindowInfo(windowHandle);

         SweepOneWindow(windowHandle, boundingBox, monitorBoundsSource, monitorBoundsTarget);
         sweptWindowsCount += 1;
      }

      if (IsFocusModeRequested())
      {
         // restore focus from swept windows (which got focus because of pushing on top for Z-reordering)
         WindowUtils.FocusOnWindow(GetCachedActiveWindow());
      }

      // save the list of windows that were swept (with position info)
      BFS.ScriptSettings.WriteValue(SweptWindowsListSetting, GetSavedWindowInfoStr());

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

            if (DetectedInvalidWindow(new IntPtr(windowHandleValue), "GetListOfWindowsToRestore", $"string window: {window}\n\nwindowsToReviveStrings: {windowsToReviveStrings}")) continue;

            windowsToRestore.Add(new IntPtr(windowHandleValue));
         }
      }

      return windowsToRestore.ToArray();
   }

   public static IntPtr[] GetListOfWindowsToUnsweep(bool forceReviveAll)
   {
      IntPtr[] listOfWindowsToUnsweep = new IntPtr[0];

      if (forceReviveAll) // "unsweep" all windows from sweep-target monitor 
      {
         // forceRevive for sweep is moving all windows from sweep-target to OLED regardless whether they were ever swept or not
         listOfWindowsToUnsweep = GetListOfWindowsToHide(GetSweepTargetMonitorID());
      }
      else // only unsweep windows previously swept
      {
         unsweepWindowsInfoMap.Clear();

         string savedWindows = BFS.ScriptSettings.ReadValue(SweptWindowsListSetting);

         // parse windows info stored in settings as string: handle,x,y,w,h,isMaximized
         var entries = savedWindows.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                           .Select(line =>
                           {
                              var parts = line.Split(',');
                              if (parts.Count() < 6)
                              {
                                 MessageBox.Show($"ERROR <min> parts < 6\nline: {line}");
                                 return (IntPtr.Zero, new Rectangle { }, false);
                              }
                              IntPtr windowHandle = (IntPtr)long.Parse(parts[0]);
                              Rectangle boundsWindow = new Rectangle(
                                 int.Parse(parts[1]), int.Parse(parts[2]),
                                 int.Parse(parts[3]), int.Parse(parts[4])
                             );
                              bool isMaximized = bool.Parse(parts[5]);

                              return (windowHandle, boundsWindow, isMaximized);
                           });

         // Store in the static map
         foreach (var entry in entries)
         {
            var (windowHandle, boundsWindow, isMaximized) = entry;
            // check if saved window handle is still valid (e.g. window wasn't closed in the meantime)
            if (!WindowUtils.IsWindowValid(windowHandle))
            {
               if (debugWindowFiltering) MessageBox.Show($"skipping invalid:\nwindowHandle {windowHandle}\nboundsWindow {boundsWindow}\nisMaximized {isMaximized}");
               continue;
            }
            if (BFS.Window.IsMinimized(windowHandle))
            {
               if (debugWindowFiltering) MessageBox.Show($"skipping saved but minimized :\nwindowHandle {windowHandle}\nboundsWindow {boundsWindow}\nisMaximized {isMaximized}");
               continue;
            }
            unsweepWindowsInfoMap[windowHandle] = (boundsWindow, isMaximized);
         }
         listOfWindowsToUnsweep = unsweepWindowsInfoMap.Keys.Cast<IntPtr>().ToArray();
      }

      listOfWindowsToUnsweepStr = string.Join("\n",
                  listOfWindowsToUnsweep.Select(window =>
                     $"{window}: |{BFS.Window.GetText(window)}| /{BFS.Window.GetClass(window)}/"));

      return listOfWindowsToUnsweep;
   }

   public static int RestoreWindows(IntPtr[] windowsToRestore)
   {
      int restoredWindowsCount = 0;
      // loop through each window to restore
      foreach (IntPtr windowHandle in windowsToRestore)
      {
         if (DetectedInvalidWindow(windowHandle, "RestoreWindows", $"minimized? {BFS.Window.IsMinimized(windowHandle)}")) continue;

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

      return restoredWindowsCount;
   }

   public static int UnsweepWindows(IntPtr[] windowsToUnsweep)
   {
      // Lazy version because those variables are common for all windows without saved position but it should not be calculated
      // in case no window needs it (all have saved positions)
      // get bounds of monitors, this time, because of "unsweeping", OLED is target and sweep-target-monitor is source 
      Lazy<Rectangle> monitorBoundsTarget = new(() => BFS.Monitor.GetMonitorBoundsByID(GetOledMonitorID()));
      Lazy<Rectangle> monitorBoundsSource = new(() => BFS.Monitor.GetMonitorBoundsByID(GetSweepTargetMonitorID()));
      Lazy<Rectangle> boundingBox = new(() => CalculateBoundingBox(windowsToUnsweep));

      int unsweptWindowsCount = 0;
      // loop through each window to restore
      foreach (IntPtr windowHandle in windowsToUnsweep)
      {
         if (debugPrintHideRevive) MessageBox.Show($"unsweeping window {BFS.Window.GetText(new IntPtr(windowHandle))}");

         // decide if position the window should be restored to saved value or swept in reverse direction
         if (HasSavedWindowInfo(windowHandle))
         {
            var savedInfo = GetSavedWindowInfo(windowHandle);

            if (BFS.Window.IsMaximized(windowHandle)) BFS.Window.Restore(windowHandle);

            WindowUtils.SetSizeAndLocation(windowHandle, savedInfo.savedPosition);

            if (savedInfo.shouldMaximize)
            {
               WindowUtils.MaximizeWindow(windowHandle, 20);
            }
         }
         else // sweeping in reverse direction because no saved window position stored
         {
            SweepOneWindow(windowHandle, boundingBox.Value, monitorBoundsSource.Value, monitorBoundsTarget.Value);
         }
         unsweptWindowsCount += 1;
      }

      // clear the windows that we saved
      BFS.ScriptSettings.WriteValue(SweptWindowsListSetting, string.Empty);

      return unsweptWindowsCount;
   }

   public static void FixZorderAfterReviveInFocusMode(IntPtr[] windowsToPushOnTop)
   {
      if (IsFocusModeRequested() && windowsToPushOnTop.Length > 0)
      {
         Array.Reverse(windowsToPushOnTop);
         foreach (IntPtr windowHandle in windowsToPushOnTop)
         {
            if (DetectedInvalidWindow(windowHandle, "FixZorderAfterReviveInFocusMode", $"windowsToPushOnTop.Length: {windowsToPushOnTop.Length}")) return;

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

   public static void SweepOneWindow(IntPtr windowHandle, Rectangle boundingBox, Rectangle boundsFrom, Rectangle boundsTo)
   {
      Rectangle boundsWindow = WindowUtils.GetBounds(windowHandle);
      bool boundsFromDoesntContain = !boundsFrom.Contains(boundsWindow);
      bool boundingBoxDoesntContain = !boundingBox.Contains(boundsWindow);

      string debugInfo = $"boundsWindow: {boundsWindow}\nboundingBox: {boundingBox}\nboundsFrom: {boundsFrom}\nboundsTo {boundsTo}\n\n" +
                         $"boundsFromDoesntContain: {boundsFromDoesntContain}\nboundingBoxDoesntContain: {boundingBoxDoesntContain}";
      if (debugPrintSweepMode) MessageBox.Show($"SweepOneWindow:\n" + debugInfo);
      if (boundsFromDoesntContain || boundingBoxDoesntContain) MessageBox.Show($"ERROR <min> SweepOneWindow:\n" + debugInfo);

      if (DetectedInvalidWindow(windowHandle, "SweepOneWindow", debugInfo)) return;

      // when window is already maximized restore it (so that it can be maximized on target screen)
      // and treat it as if its size would be the same as source monitors size
      // (because GetBounds() function gives you window bounds from before maximizing it)
      if (BFS.Window.IsMaximized(windowHandle))
      {
         BFS.Window.Restore(windowHandle);
         boundsWindow = boundsFrom;
      }

      var result = CalculateSweptWindowPos(boundsWindow, boundingBox, boundsFrom, boundsTo);

      WindowUtils.SetSizeAndLocation(windowHandle, result.newBoundsWindow);
      WindowUtils.PushToTop(windowHandle);

      // maximize window when it is big enough
      if (result.shouldMaximize)
      {
         WindowUtils.MaximizeWindow(windowHandle, 20);
      }
   }

   public static (bool shouldMaximize, Rectangle newBoundsWindow) CalculateSweptWindowPos(
                                                                           Rectangle boundingBox,
                                                                           Rectangle boundsWindow,
                                                                           Rectangle boundsFrom,
                                                                           Rectangle boundsTo)
   {
      if (boundingBox.Width <= RESOLUTION_2K_WIDTH && boundingBox.Height <= RESOLUTION_2K_HEIGHT) // or SWEEP_NO_RESIZE_THRESHOLD??
      {
         return CalculateSweptWindowPosOneToOne(boundsWindow, boundingBox, boundsFrom, boundsTo);

      }

      return CalculateSweptWindowPosDynamic(boundsWindow, boundingBox, boundsFrom, boundsTo);
   }

   public static (bool shouldMaximize, Rectangle newBoundsWindow) CalculateSweptWindowPosOneToOne(
                                                                           Rectangle boundingBox,
                                                                           Rectangle boundsWindow,
                                                                           Rectangle boundsFrom,
                                                                           Rectangle boundsTo)
   {
      // Rectangle newBoundsWindow = new Rectangle();
      // bool shouldMaximize = false;

      // return (shouldMaximize, newBoundsWindow);
      return CalculateSweptWindowPosDynamic(boundingBox, boundsWindow, boundsFrom, boundsTo); // todo
   }
   public static (bool shouldMaximize, Rectangle newBoundsWindow) CalculateSweptWindowPosDynamic(
                                                                           Rectangle boundingBox,
                                                                           Rectangle boundsWindow,
                                                                           Rectangle boundsFrom,
                                                                           Rectangle boundsTo)
   {
      Rectangle newBoundsWindow = new Rectangle();
      bool shouldMaximize = ShouldMaximizeWindow(boundsWindow, boundsTo);

      if (shouldMaximize)
      {
         // set size to max bounds below which windows are not resized and position to center
         // this will set the size to which window will restore when unmaximizing
         newBoundsWindow = GetWindowPosNoResizeThreshold(boundsTo);
         shouldMaximize = true;

         if (debugPrintSweepModeCalcPos) MessageBox.Show($"\nboundsWindow\t{boundsWindow}\n" +
                                             $"boundsFrom\t{boundsFrom}\n" +
                                             $"boundsTo\t{boundsTo}\n\n" +
                                             $"newBoundsWindow\t{newBoundsWindow}\n" +
                                             $"noResizeMaxWidth\t{GetNoResizeMaxWidth(boundsTo)}\t" +
                                             $"maxHeight\t{GetNoResizeMaxHeight(boundsTo)}\n" +
                                             $"shouldMaximize:\t{shouldMaximize}");
      }
      else // should NOT maximize
      {
         // small window in at least one direction, don't change size for now
         newBoundsWindow.Width = boundsWindow.Width;
         newBoundsWindow.Height = boundsWindow.Height;

         // calculate how far is the window from old monitor borders (ignore taskbar size because of autohide)
         var borderDistances = CalculateBorderDistances(boundsWindow, boundsFrom);

         // calculate left-right (horizontal) and top-bottom (vertical) ratios of window position
         var ratios = CalculateRatios(borderDistances.left, borderDistances.right, borderDistances.top, borderDistances.bottom);

         // calculate how much free space there will be on new monitor outside of the window
         var freeSpace = CalculateSpaceToBorders(boundsTo, newBoundsWindow.Width, newBoundsWindow.Height);

         if (freeSpace.horizontal > 0)
         {
            // set position to preserve left-right ratio from old monitor
            newBoundsWindow.X = boundsTo.X + (int)(ratios.horizontal * freeSpace.horizontal);
         }
         else // freeSpace.horizontal <= 0
         {
            // there is no space left, max out window horizontally
            newBoundsWindow.X = boundsTo.X;
            // override window witdth, it needs to be smaller
            newBoundsWindow.Width = boundsTo.Width;

            // optionally snap top or bottom if window was resized horizontally
            if (enableSweepModeSnap)
            {
               int topDistNew = newBoundsWindow.Y - boundsTo.Y;
               int bottomDistNew = (boundsTo.Y + boundsTo.Height - TASKBAR_HEIGHT) - (newBoundsWindow.Y + newBoundsWindow.Height);

               if (topDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap top");
                  newBoundsWindow.Y = 0; // snap window to top, no size change
               }

               if (bottomDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap bottom");
                  newBoundsWindow.Y = boundsTo.Height - TASKBAR_HEIGHT - newBoundsWindow.Height; // snap window to bottom, no size change
               }

               if (debugPrintSweepModeCalcPos) MessageBox.Show($"topDistNew {topDistNew}\nbottomDistNew {bottomDistNew}\n\n" +
                                                $"newBoundsWindow.Y {newBoundsWindow.Y}");
            }

         }

         if (freeSpace.vertical > 0)
         {
            // set position to preserve top-bottom ratio from old monitor
            newBoundsWindow.Y = boundsTo.Y + (int)(ratios.vertical * freeSpace.vertical);
         }
         else // freeSpace.vertical <= 0
         {
            // there is no space left, max out window vertically
            newBoundsWindow.Y = boundsTo.Y;
            // override window height, it needs to be smaller
            newBoundsWindow.Height = boundsTo.Height - TASKBAR_HEIGHT;

            // optionally snap left or right if window was resized vertically
            if (enableSweepModeSnap)
            {
               int leftDistNew = newBoundsWindow.X - boundsTo.X;
               int rightDistNew = (boundsTo.X + boundsTo.Width) - (newBoundsWindow.X + newBoundsWindow.Width);

               // check if old window width was near to half split of the source monitor
               bool widthEligibleForHalfSplit = boundsWindow.Width < (boundsFrom.Width / 2 + SWEEP_SNAP_HALFSPLIT_THRESHOLD) &&
                                                boundsWindow.Width > (boundsFrom.Width / 2 - SWEEP_SNAP_HALFSPLIT_THRESHOLD);

               if (leftDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap left");
                  newBoundsWindow.X = 0; // snap window to left, no size change

                  if (enableSweepModeSnapHalfSplit && widthEligibleForHalfSplit)
                  {
                     if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap to vertical half split on right");
                     newBoundsWindow.Width = boundsTo.Width / 2;
                  }
               }

               if (rightDistNew < SWEEP_SNAP_THRESHOLD)
               {
                  if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap right");

                  if (enableSweepModeSnapHalfSplit && widthEligibleForHalfSplit)
                  {
                     if (debugPrintSweepModeCalcPos) MessageBox.Show($"Snap to vertical half split on left");
                     newBoundsWindow.Width = boundsTo.Width / 2;
                  }

                  newBoundsWindow.X = boundsTo.Width - newBoundsWindow.Width; // snap window to right after optional size change
               }

               if (debugPrintSweepModeCalcPos) MessageBox.Show($"leftDistNew {leftDistNew}\nrightDistNew {rightDistNew}\n\n" +
                                                               $"newBoundsWindow.X {newBoundsWindow.X}");
            }
         }

         bool sizeWarning = newBoundsWindow.Width < 20 || newBoundsWindow.Width > 3840 ||
                        newBoundsWindow.Height < 20 || newBoundsWindow.Height > 2160 ||
                        newBoundsWindow.X < -3840 || newBoundsWindow.X > 2160 ||
                        newBoundsWindow.Y < -2160 || newBoundsWindow.Y > (2160 + 1920);
         bool doesntContain = !boundsFrom.Contains(boundsWindow);

         if (debugPrintSweepModeCalcPos || sizeWarning || doesntContain) MessageBox.Show($"sizeWarning? {sizeWarning}\tdoesntContain? {doesntContain}\n" +
                         $"listOfWindowsToUnsweep {listOfWindowsToUnsweepStr}\n" +
                         $"listOfWindowsToHide {listOfWindowsToHideStr}\n\n" +
                         $"boundsWindow\t{boundsWindow}\nboundsFrom\t{boundsFrom}\nboundsTo\t{boundsTo}\n\n" +
                         $"borderDistances.left {borderDistances.left}\nborderDistances.right {borderDistances.right}\n" +
                         $"borderDistances.top {borderDistances.top}\nborderDistances.bottom {borderDistances.bottom}\n\n" +
                         $"horizontalRatio {ratios.horizontal}\nverticalRatio {ratios.vertical}\n\n" +
                         $"horizontalSpace {freeSpace.horizontal}\nverticalSpace {freeSpace.vertical}\n\n" +
                         $"newBoundsWindow.X {newBoundsWindow.X}\nnewBoundsWindow.Y {newBoundsWindow.Y}\n\n" +
                         $"newBoundsWindow.Width {newBoundsWindow.Width}\nnewBoundsWindow.Height {newBoundsWindow.Height}");
      }


      return (shouldMaximize, newBoundsWindow);
   }

   public static Rectangle GetWindowPosNoResizeThreshold(Rectangle boundsTo)
   {
      int noResizeMaxWidth = GetNoResizeMaxWidth(boundsTo);
      int noResizeMaxHeight = GetNoResizeMaxHeight(boundsTo);

      return new Rectangle(
         boundsTo.X + (boundsTo.Width - noResizeMaxWidth) / 2,
         boundsTo.Y + (boundsTo.Height - noResizeMaxHeight) / 2,
         noResizeMaxWidth,
         noResizeMaxHeight
      );
   }

   public static int GetNoResizeMaxWidth(Rectangle boundsTo)
   {
      return (int)(SWEEP_NO_RESIZE_THRESHOLD * boundsTo.Width);
   }
   public static int GetNoResizeMaxHeight(Rectangle boundsTo, int taskbar = TASKBAR_HEIGHT)
   {
      return (int)(SWEEP_NO_RESIZE_THRESHOLD * (boundsTo.Height - taskbar));
   }
   public static bool ShouldMaximizeWindow(Rectangle boundsWindow, Rectangle boundsTo)
   {
      return (boundsWindow.Width >= GetNoResizeMaxWidth(boundsTo)) &&
             (boundsWindow.Height >= GetNoResizeMaxHeight(boundsTo));
   }

   public static (int left, int right, int top, int bottom) CalculateBorderDistances(Rectangle boundsWindow,
                                                                                     Rectangle boundsMonitor)
   {
      if (!boundsMonitor.Contains(boundsWindow))
         MessageBox.Show($"ERROR <min> CalculateBorderDistances:\nboundsFrom {boundsMonitor}\ndoesn't contain\nboundsWindow {boundsWindow}");

      int left = ClipBottom(boundsWindow.X - boundsMonitor.X); // +1
      int right = ClipBottom((boundsMonitor.X + boundsMonitor.Width) - (boundsWindow.X + boundsWindow.Width)); // +1
      int top = ClipBottom(boundsWindow.Y - boundsMonitor.Y); // +1
      int bottom = ClipBottom((boundsMonitor.Y + boundsMonitor.Height) - (boundsWindow.Y + boundsWindow.Height)); // +1

      return (left, right, top, bottom);
   }

   public static (double horizontal, double vertical) CalculateRatios(int left, int right, int top, int bottom)
   {
      double horizontal = 0.5;
      if ((left + right) == 0)
      {
         MessageBox.Show($"WARN <min> CalculateRatios (left + right) == 0");
      }
      else
      {
         horizontal = (double)left / (left + right);
      }

      double vertical = 0.5;
      if ((top + bottom) == 0)
      {
         MessageBox.Show($"WARN <min> CalculateRatios (top + bottom) == 0");
      }
      else
      {
         vertical = (double)top / (top + bottom);
      }

      return (horizontal, vertical);
   }

   public static (int horizontal, int vertical) CalculateSpaceToBorders(Rectangle boundsMonitor, int windowWitdth, int windowHeight)
   {
      int horizontal = boundsMonitor.Width - windowWitdth;
      int vertical = boundsMonitor.Height - windowHeight - TASKBAR_HEIGHT;

      return (horizontal, vertical);
   }

   public static Rectangle CalculateBoundingBox(IntPtr[] windowHandles)
   {
      if (windowHandles.Length == 0)
      {
         MessageBox.Show($"ERROR <min>! CalculateBoundingBox when no windows on list");
         return new Rectangle { };
      }

      if (!enableBoundingBoxMode)
      {
         // if bounding box mode is disabled return monitor bounds as bounding box always
         return BFS.Monitor.GetMonitorBoundsByWindow(windowHandles[0]);
      }

      Rectangle boundingBox = WindowUtils.GetBounds(windowHandles[0]);
      foreach (IntPtr windowHandle in windowHandles)
      {
         Rectangle boundsWindow = WindowUtils.GetBounds(windowHandle);
         boundingBox = Rectangle.Union(boundingBox, boundsWindow);

         bool doesntContain = !boundingBox.Contains(boundsWindow);
         if (doesntContain) MessageBox.Show($"CalculateBoundingBox \n\n" +
                         $"|{BFS.Window.GetText(windowHandle)}|\n\n" +
                         $"doesntContain? {doesntContain} \n\n" +
                         $"boundsWindow {boundsWindow}\nnot in\nboundingBox {boundingBox}");
      }

      return boundingBox;
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

      MessageBox.Show($"ERROR <min>! did not find monitor with 4K resolution");
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

      MessageBox.Show($"ERROR <min>! did not find monitor with 2K resolution");
      return UInt32.MaxValue;
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

   public static void Init()
   {
      CacheAll();

      listOfWindowsToUnsweepStr = "";
      listOfWindowsToHideStr = "";
      unsweepWindowsInfoMap.Clear();
      saveInfoStrBuilder.Clear();
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
      ActiveWindowHandleCache = BFS.Window.GetFocusedWindow();
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
      return ActiveWindowHandleCache;
   }

   public static IntPtr[] GetFilteredVisibleWindows(uint monitorId)
   {
      return BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId).Where(windowHandle =>
      {
         if (IsWindowBlacklisted(windowHandle))
         {
            if (debugWindowFiltering) MessageBox.Show($"W ignored blacklisted:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
            return false;
         }
         else
         {
            if (debugWindowFiltering) MessageBox.Show($"W NOT ignored blacklisted:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         }

         if (BFS.Window.IsMinimized(windowHandle))
         {
            if (debugWindowFiltering) MessageBox.Show($"W ignored minimized:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
            return false;
         }
         else
         {
            if (debugWindowFiltering) MessageBox.Show($"W NOT ignored minimized:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         }

         if (debugWindowFiltering) MessageBox.Show($"W NOT ignored at all:\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         return true;
      }).ToArray();
   }

   public static IntPtr[] GetListOfWindowsToHide(uint monitorId)
   {
      // generate debug list and filter out focused window in active mode
      var filteredWindows = GetFilteredVisibleWindows(monitorId)
          .Select(windowHandle => new
          {
             Handle = windowHandle,
             Text = BFS.Window.GetText(windowHandle),
             Class = BFS.Window.GetClass(windowHandle)
          })
          .Where(window =>
          {
             if (IsFocusModeRequested() && window.Handle == GetCachedActiveWindow())
             {
                //  MessageBox.Show($"Ignored focused window as window to hide: \n\n" +
                //                  $"|{window.Text}|\n\n|{window.Class}|");
                return false;
             }
             return true;
          });

      // save debug info
      listOfWindowsToHideStr = string.Join("\n",
          filteredWindows.Select(w => $"{w.Handle}: |{w.Text}| /{w.Class}/"));

      // return array of window handles only
      return filteredWindows.Select(w => w.Handle).ToArray();
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
      // ignore window handles that no longer exists
      if (!WindowUtils.IsWindowValid(windowHandle))
      {
         MessageBox.Show($"ignoring not valid window: {windowHandle}\n\n|{BFS.Window.GetText(windowHandle)}|\n\n|{BFS.Window.GetClass(windowHandle)}|");
         return true; // todo comment out print
      }

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
         if (debugWindowFiltering) MessageBox.Show($"Filtered out windows wrong size (w{windowRect.Width}, h{windowRect.Height}). classname: {classname}, text: {text})");
         return true;
      }

      return false;
   }

   public static (int X, int Y) GetMouseHideTarget()
   {
      Rectangle bounds = BFS.Monitor.GetPrimaryMonitorBounds();
      return (bounds.Width / 2, bounds.Height / 2);
   }

   public static bool DetectedInvalidWindow(IntPtr windowHandle, string windowName, string debugInfo)
   {
      if (!WindowUtils.IsWindowValid(windowHandle))
      {
         MessageBox.Show($"ERROR <min>! {windowName}\n" +
                        $" detected invalid window handle:\n{windowHandle} |{BFS.Window.GetText(windowHandle)}|\n\n" +
                        $"listOfWindowsToUnsweep {listOfWindowsToUnsweepStr}\n" +
                        $"listOfWindowsToHide {listOfWindowsToHideStr}\n\n" +
                        $"{debugInfo}");
         return true;
      }
      return false;
   }

   public static int ClipBottom(int val)
   {
      if (val < 0) MessageBox.Show($"WARN <min> Clipping bottom: {val}");
      return val < 0 ? 0 : val;
   }

   private static void AppendSavedWindowInfo(IntPtr windowHandle)
   {
      Rectangle bounds = WindowUtils.GetBounds(windowHandle);
      bool isMaximized = BFS.Window.IsMaximized(windowHandle);
      saveInfoStrBuilder.AppendLine($"{windowHandle},{bounds.X},{bounds.Y},{bounds.Width},{bounds.Height},{isMaximized}");
   }

   public static string GetSavedWindowInfoStr()
   {
      string str = saveInfoStrBuilder.ToString();
      saveInfoStrBuilder.Clear();
      return str;
   }

   public static (Rectangle savedPosition, bool shouldMaximize) GetSavedWindowInfo(IntPtr windowHandle)
   {
      if (unsweepWindowsInfoMap.Contains(windowHandle))
      {
         return ((Rectangle, bool))unsweepWindowsInfoMap[windowHandle]!;
      }

      MessageBox.Show("ERROR <min>! GetSavedWindowInfo: window handle {windowHandle} not found in the map");

      return (new Rectangle { }, false);
   }


   private static bool HasSavedWindowInfo(IntPtr windowHandle)
   {
      return unsweepWindowsInfoMap.Contains(windowHandle);
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
            MessageBox.Show($"ERROR <min> window not valid in GetBounds: {windowHandle}");
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
                          $"new pos: {newPos} (shad-comp)\n" +
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
            MessageBox.Show($"ERROR <min> window not valid in PushToTop: {windowHandle}");
            return;
         }

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