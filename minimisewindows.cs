using System;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;

using System.Text;
using System.Collections.Generic;
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
    public const bool logStartTimerDefault = false;
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
    private const bool debugPrintCalculateSnapPointLen = enableDebugPrints && false;

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
        using (Log.T("whole program", Log.LogLevel.Info, startLog: false))
        {
            using (Log.T("Initialization", startLog: false)) { Init(); }

            IntPtr[] windowsToHide;
            using (Log.T("GetListOfWindowsToHide")) { windowsToHide = GetListOfWindowsToHide(GetOledMonitorID()); }

            bool windowsToHidePresent = windowsToHide.Length > 0;
            bool windowsWereMinimized = WereWindowsMinimized();
            bool windowsWereSwept = WereWindowsSwept();
            bool windowsWereHiddenCurrentMode = (!IsSweepModeRequested() && windowsWereMinimized) ||
                                                (IsSweepModeRequested() && windowsWereSwept);

            Log.D("Initial state", new { windowsToHide, windowsToHidePresent, windowsWereMinimized, windowsWereSwept, windowsWereHiddenCurrentMode });

            if (IsForceReviveRequested())
            {
                using (Log.T("Forced HandleReviving", startLog: true)) { HandleReviving(false); }
            }
            else // no forceRevive
            {
                if (windowsToHidePresent && !windowsWereHiddenCurrentMode)
                {
                    using (Log.T("HandleHiding present windows", startLog: true)) { HandleHiding(windowsToHide); }
                }
                else if (windowsToHidePresent && windowsWereHiddenCurrentMode)
                {
                    if (ShouldPrioritizeHiding())
                    {
                        // ignore saved windows and hide visible ones
                        using (Log.T("HandleHiding PrioritizeHiding", startLog: true)) { HandleHiding(windowsToHide); }
                    }
                    else
                    {
                        // ignore visible windows and revive saved ones
                        using (Log.T("HandleReviving with no PrioritizeHiding", startLog: true)) { HandleReviving(false); }
                    }
                }
                else if (!windowsToHidePresent && windowsWereHiddenCurrentMode)
                {
                    // revive saved windows only
                    using (Log.T("HandleReviving saved windows", startLog: true)) { HandleReviving(false); }
                }
                else if (!windowsToHidePresent && !windowsWereHiddenCurrentMode)
                {
                    Log.D("NOTHING to hide and NOT hidden. Doing Nothing!");
                    // HandleReviving(true); // todo force revive all windows or do nothing?
                }
                else
                {
                    Log.E("else state not expected!", showMessageBox: true);
                }
            }

            Log.D("program finished", new { listOfWindowsToUnsweepStr, listOfWindowsToHideStr });
        } // timed operation whole program

        // todo Logger.Warnings / Errors present message box
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
            Log.I($"minimizing window {BFS.Window.GetText(windowHandle)}");

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

        Log.I($"finished MIN (minimized {minimizedWindowsCount}/{windowsToMinimize.Length} windows)");

        return minimizedWindowsCount;
    }

    public static int SweepWindows(IntPtr[] windowsToSweep)
    {
        if (windowsToSweep.Length == 0)
        {
            Log.E("no windows on sweep list");
            return 0;
        }

        // get bounds of monitors
        Rectangle monitorBoundsTarget = BFS.Monitor.GetMonitorBoundsByID(GetSweepTargetMonitorID());
        Rectangle monitorBoundsSource = BFS.Monitor.GetMonitorBoundsByID(GetOledMonitorID());

        // sweep mode in reverse order (the same as restoring) compared to minimizing windows
        Array.Reverse(windowsToSweep);

        // calculate bounding box of all windows
        Rectangle boundingBox = CalculateBoundingBox(windowsToSweep);

        Log.I("starting sweeping windows", new { windowsToSweep, boundingBox, monitorBoundsSource, monitorBoundsTarget });

        // loop through all windows to be swept
        int sweptWindowsCount = 0;
        foreach (IntPtr windowHandle in windowsToSweep)
        {
            Log.I("sweeping window", () => new { windowHandle, text = BFS.Window.GetText(windowHandle), sweptWindowsCount });

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

        Log.I($"finished sweep", new { sweptWindowsCount, windowsToSweep.Length });

        return sweptWindowsCount;
    }

    public static void HandleReviving(bool reviveAll)
    {
        Log.I("start reviving windows", new { reviveAll });

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

        Log.I($"finished revive", new { count, windowsToRevive.Length, keepReviving });
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
                                      Log.E("parts < 6", new { line }, showMessageBox: true);
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
                    Log.I("skipping invalid window", new { windowHandle, boundsWindow, isMaximized });
                    continue;
                }
                if (BFS.Window.IsMinimized(windowHandle))
                {
                    Log.I("skipping saved but minimized window", new { windowHandle, boundsWindow, isMaximized });
                    continue;
                }
                unsweepWindowsInfoMap[windowHandle] = (boundsWindow, isMaximized);
            }
            listOfWindowsToUnsweep = unsweepWindowsInfoMap.Keys.Cast<IntPtr>().ToArray();
            Log.D("got list of windows to unsweep (map)", new { unsweepWindowsInfoMap });
            Log.D("got list of windows to unsweep (keys)", new { listOfWindowsToUnsweep });
        }

        Log.D("got list of windows to unsweep (with text and class)", () => new
        {
            windows = listOfWindowsToUnsweep.Select(window => $"{window}: |{BFS.Window.GetText(window)}| /{BFS.Window.GetClass(window)}/")
        });
        // todo remove listOfWindowsToUnsweepStr
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
                Log.I($"restoring window", () => new { text = BFS.Window.GetText(windowHandle) });
                WindowUtils.RestoreWindow(windowHandle);
                WindowUtils.PushToTop(windowHandle);
                restoredWindowsCount += 1;
            }
            else
            {
                Log.D($"already restored window", () => new { text = BFS.Window.GetText(windowHandle) });
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
            Log.I($"unsweeping window", () => new { text = BFS.Window.GetText(windowHandle) });

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

                Log.D($"Pushing window on top (focus mode)", () => new { text = BFS.Window.GetText(windowHandle) });
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
            Log.I($"hiding mouse", new { mouseX, mouseY });
        }
        else // mouse on other monitor
        {
            // clear stored mouse position
            BFS.ScriptSettings.DeleteValue(MousePositionXSetting);
            BFS.ScriptSettings.DeleteValue(MousePositionYSetting);
            Log.D($"skip hiding mouse because not on 4k OLED monitor", new { mouseMonitorBounds, mouseX, mouseY });
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
            Log.D("no stored mouse position");
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

        Log.D("mouse debug info", new { mouseOldX, mouseOldY, mouseX, mouseY, mouseHideTargetX, mouseHideTargetY, diffX, diffY, wasMoved });

        if (!wasMoved)
        {
            // restore mouse position because it wasn't moved enough
            Log.I($"restoring mouse position to saved coordinates");
            BFS.Input.SetMousePosition(mouseOldX, mouseOldY);
        }
        else
        {
            Log.D($"skiping restoring mouse was moved too much");
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

        // when window is already maximized restore it before sweeping (so that it can be maximized on target screen)
        if (BFS.Window.IsMaximized(windowHandle))
        {
            BFS.Window.Restore(windowHandle);
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
        if (boundingBox.Width <= RESOLUTION_2K_WIDTH && boundingBox.Height <= RESOLUTION_2K_HEIGHT)
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
                // there is no space left horizontally, max out window left-right
                newBoundsWindow.X = boundsTo.X;
                newBoundsWindow.Width = boundsTo.Width;

                // optionally snap top or bottom if window was resized horizontally

                var snapped = CalculateSnapPointLen(newBoundsWindow.Y, newBoundsWindow.Height,
                                                    boundsTo.Y, boundsTo.Height - TASKBAR_HEIGHT, false);
                newBoundsWindow.Y = snapped.point;
                newBoundsWindow.Height = snapped.len;
            }

            if (freeSpace.vertical > 0)
            {
                // set position to preserve top-bottom ratio from old monitor
                newBoundsWindow.Y = boundsTo.Y + (int)(ratios.vertical * freeSpace.vertical);
            }
            else // freeSpace.vertical <= 0
            {
                // there is no space left vertically, max out window top-bottom
                newBoundsWindow.Y = boundsTo.Y;
                newBoundsWindow.Height = boundsTo.Height - TASKBAR_HEIGHT;

                bool widthEligibleForHalfSplit = boundsWindow.Width < (boundsFrom.Width / 2 + SWEEP_SNAP_HALFSPLIT_THRESHOLD) &&
                                     boundsWindow.Width > (boundsFrom.Width / 2 - SWEEP_SNAP_HALFSPLIT_THRESHOLD);

                // optionally snap left or right if window was resized vertically
                var snapped = CalculateSnapPointLen(newBoundsWindow.X, newBoundsWindow.Width,
                                                    boundsTo.X, boundsTo.Width,
                                                    widthEligibleForHalfSplit);
                newBoundsWindow.X = snapped.point;
                newBoundsWindow.Width = snapped.len;

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

    public static (int point, int len) CalculateSnapPointLen(int windowPoint, int windowLen, int monitorPoint, int monitorLen, bool doHalfSplit)
    {
        // optionally snap left or right if window was resized vertically
        int newWindowPoint = windowPoint;
        int newWindowLen = windowLen;

        if (enableSweepModeSnap)
        {
            int leftTopDistNew = windowPoint - monitorPoint;
            int rightBottomDistNew = (monitorPoint + monitorLen) - (windowPoint + windowLen);

            string debugInfo = $"windowPoint: {windowPoint}\twindowLen: {windowLen}\n" +
                               $"monitorPoint: {monitorPoint}\tmonitorLen: {monitorLen}\n" +
                               $"doHalfSplit: {doHalfSplit}\n\n" +
                               $"leftTopDistNew: {leftTopDistNew}\n" +
                               $"rightBottomDistNew: {rightBottomDistNew}\n";
            if (leftTopDistNew < 0) MessageBox.Show($"ERR <min> CalculateSnapPointLen: leftTopDistNew < 0\n\n" + debugInfo);
            if (rightBottomDistNew < 0) MessageBox.Show($"ERR <min> CalculateSnapPointLen: rightBottomDistNew < 0\n\n" + debugInfo);

            if (leftTopDistNew < SWEEP_SNAP_THRESHOLD)
            {
                if (debugPrintCalculateSnapPointLen) MessageBox.Show($"CalculateSnapPointLen\nSnap left/top");
                newWindowPoint = 0; // snap window to left/top, no size change

                if (enableSweepModeSnapHalfSplit && doHalfSplit)
                {
                    if (debugPrintCalculateSnapPointLen) MessageBox.Show($"CalculateSnapPointLen\nSnap to half split on right/bottom");
                    newWindowLen = monitorLen / 2;
                }
            }

            if (rightBottomDistNew < SWEEP_SNAP_THRESHOLD)
            {
                if (debugPrintCalculateSnapPointLen) MessageBox.Show($"CalculateSnapPointLen\nSnap right/bottom");

                if (enableSweepModeSnapHalfSplit && doHalfSplit)
                {
                    if (debugPrintCalculateSnapPointLen) MessageBox.Show($"CalculateSnapPointLen\nSnap to vertical half split on left/top");
                    newWindowLen = monitorLen / 2;
                }

                newWindowPoint = monitorLen - newWindowLen; // snap window to right/bottom after optional size change
            }

            if (debugPrintCalculateSnapPointLen) MessageBox.Show($"CalculateSnapPointLen\n\n" +
                                                                  $"newWindowPoint {newWindowPoint}\n" +
                                                                  $"newWindowLen {newWindowLen}\n\n" +
                                                                  debugInfo);
        }

        return (newWindowPoint, newWindowLen);
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
        Log.I($"\n\n================================================================================\n" +
              $"\t\t\t\t OLED minimizer - {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
              $"================================================================================",
              skipHeader: true);
        Log.D("Initializing");
        CacheAll();

        Log.D("Clearing variables");
        listOfWindowsToUnsweepStr = "";
        listOfWindowsToHideStr = "";
        unsweepWindowsInfoMap.Clear();
        saveInfoStrBuilder.Clear();
    }

    public static void CacheAll()
    {
        CacheForceReviveRequest();
        CacheFocusModeRequest();
        CacheSweepModeRequest();
        CacheActiveWindow();
    }

    public static void CacheForceReviveRequest()
    {
        // bool ForceReviveRequestedCache = BFS.Input.IsMouseDown("1;");
        ForceReviveRequestedCache = BFS.Input.IsKeyDown(KEY_SHIFT);
        Log.I($"ForceRevive key is" + (ForceReviveRequestedCache ? "" : " NOT") + " pressed, caching");
    }

    public static void CacheFocusModeRequest()
    {
        // bool FocusModeRequestedCache = BFS.Input.IsMouseDown("2;");
        FocusModeRequestedCache = BFS.Input.IsKeyDown(KEY_CTRL);
        Log.I($"FocusMode key is" + (FocusModeRequestedCache ? "" : " NOT") + " pressed, caching");
    }

    public static void CacheSweepModeRequest()
    {
        // bool SweepModeRequestedCache = BFS.Input.IsMouseDown("2;");
        SweepModeRequestedCache = BFS.Input.IsKeyDown(KEY_ALT);
        Log.I($"SweepMode key is" + (SweepModeRequestedCache ? "" : " NOT") + " pressed, caching");
    }

    public static void CacheActiveWindow()
    {
        ActiveWindowHandleCache = BFS.Window.GetFocusedWindow();
        Log.I($"Cached focused window ", new { ActiveWindowHandleCache, text = BFS.Window.GetText(ActiveWindowHandleCache) });
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
                Log.D($"Window ignored blacklisted", () => new { text = BFS.Window.GetText(windowHandle), classname = BFS.Window.GetClass(windowHandle) });
                return false;
            }

            if (BFS.Window.IsMinimized(windowHandle))
            {
                Log.D($"Window ignored minimized", () => new { text = BFS.Window.GetText(windowHandle), classname = BFS.Window.GetClass(windowHandle) });
                return false;
            }

            Log.I($"Window NOT ignored at all", () => new { text = BFS.Window.GetText(windowHandle), classname = BFS.Window.GetClass(windowHandle) });
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
            Log.W($"ignoring not valid window", () => new { windowHandle, text = BFS.Window.GetText(windowHandle), classname = BFS.Window.GetClass(windowHandle) });
            return true;
        }

        // ignore windows based on classname blacklist
        string classname = BFS.Window.GetClass(windowHandle);
        if (classnameBlacklist.Exists(blacklistItem =>
              {
                  if (classname.StartsWith(blacklistItem, StringComparison.Ordinal))
                  {
                      Log.D($"Ignored bacause of class", new { classname, blacklistItem });
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
            Log.D($"Ignored bacause of empty text", new { classname });
            return true;
        }

        // ignore windows based on text blacklist 
        if (textBlacklist.Exists(blacklistItem =>
              {
                  if (text.Equals(blacklistItem, StringComparison.Ordinal))
                  {
                      Log.D($"Ignored bacause of text", new { text, blacklistItem });
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
            Log.W($"Ignored bacause of wrong size", new { bounds = windowRect, classname, text });
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
            // todo
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
        Log.W($"clipping bottom: {val}", condition: () => (val < 0));
        return val < 0 ? 0 : val;
    }

    private static void AppendSavedWindowInfo(IntPtr windowHandle)
    {
        bool isMaximized = BFS.Window.IsMaximized(windowHandle);
        Rectangle bounds = isMaximized ?
                       WindowUtils.GetRestoredBounds(windowHandle) :
                       WindowUtils.GetBounds(windowHandle);

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

        Log.E($"window handle {windowHandle} not found in the map", new { unsweepWindowsInfoMap });

        return (new Rectangle { }, false);
    }


    private static bool HasSavedWindowInfo(IntPtr windowHandle)
    {
        return unsweepWindowsInfoMap.Contains(windowHandle);
    }
}
