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
	
	public static IntPtr[] getFilteredWindows(uint monitorId)
	{
		IntPtr[] allWindows = BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId);
			
		IntPtr[] filteredWindows = allWindows.Where(windowHandle => {

			string classname = BFS.Window.GetClass(windowHandle);
			//Rectangle windowRect = WindowUtils.GetBounds(windowHandle);

			if(classname.StartsWith("DFTaskbar") || 
			   classname.StartsWith("DFTitleBarWindow") ||
			   classname.StartsWith("Shell_TrayWnd") || 
			   classname.StartsWith("tooltips") ||
			   classname.StartsWith("Shell_InputSwitchTopLevelWindow") || // language swich in taskbar
			   classname.StartsWith("Windows.UI.Core.CoreWindow") ||  // Start and Search windows
			   classname.StartsWith("Progman") || // Program Manager
			   classname.StartsWith("SizeTipClass") || // When resizing
               classname.StartsWith("DF", StringComparison.Ordinal) ||
               classname.Equals("WorkerW", StringComparison.Ordinal) ||
               classname.Equals("SearchPane", StringComparison.Ordinal))
			{
				return false;
			}

			if (BFS.Window.IsMinimized(windowHandle)) // ignore minimized
			{
				return false;
			}

			string text = BFS.Window.GetText(windowHandle);
			if (string.IsNullOrEmpty(text) 
			    || text == "Program Manager" // also class Progman
			    || text == "Volume Mixer" // can be moved but prefer not to
			    || text == "Snap Assist"
			    || text == "Greenshot capture form" // when selecting area to screenshot (also maybe can be filtered out by size of all monitors)
			    || text == "Battery Information" // also class Windows.UI.Core.CoreWindow
			    || text == "Date and Time Information" // also class Windows.UI.Core.CoreWindow
			    || text == "Network Connections" // also class Windows.UI.Core.CoreWindow
			    || text == "Volume Control" // also class Windows.UI.Core.CoreWindow
			    || text == "Start" // also class Windows.UI.Core.CoreWindow
			    || text == "Search" // also class Windows.UI.Core.CoreWindow
				)
			{
				return false;
			}

			
			//if (windowRect.Width <= 0 || windowRect.Height <= 0)
			//{
			//	return false;
			//}

			return true;
		}).ToArray();

		return filteredWindows;
	}

    public static uint GetOledMonitorID()
    {
        foreach(uint id in BFS.Monitor.GetMonitorIDs())
        {
            Rectangle bounds = BFS.Monitor.GetMonitorBoundsByID(id);
            if(bounds.Width == RESOLUTION_4K_WIDTH && bounds.Height == RESOLUTION_4K_HEIGHT) 
            {
                if(debugPrintFindMonitorId) MessageBox.Show($"found 4k monitor with ID: {id}");            
                return id;
            }
        }
        
        MessageBox.Show($"ERROR! did not find monitor with 4k resolution");            
        return UInt32.MaxValue;
    }
	
	public static void Run(IntPtr windowHandle)
	{
		// check to see if we are minimizing 
		if(IsScriptInMinimizeState())
		{
			if(debugPrintStartStop) MessageBox.Show("start MIN");
			// this will store the windows that we are minimizing so we can restore them later
			string minimizedWindows = "";
			
			// get the monitor that the cursor is on
			// uint monitorId = BFS.Monitor.GetMonitorIDByXY(BFS.Input.GetMousePositionX(), BFS.Input.GetMousePositionY());
           
            // get monitor ID of OLED monitor (assumption it is the only 4k monitor in the system)
            uint monitorId = GetOledMonitorID();


			// loop through all the visible windows on the cursor monitor
			foreach(IntPtr window in getFilteredWindows(monitorId))
			{
				// minimize the window
				if(debugPrint) MessageBox.Show($"minimizing {BFS.Window.GetText(window)}");
				BFS.Window.Minimize(window);
				
				// add the window to the list of windows
				minimizedWindows += window.ToInt64().ToString() + "|";
			}
			
			// save the list of windows we minimized
			BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, minimizedWindows);
			
			// set the script state to NormalizeState
			BFS.ScriptSettings.WriteValue(ScriptStateSetting, NormalizeState);
			
			if(debugPrintStartStop) MessageBox.Show("finished MIN");
		}
		else
		{
			if(debugPrintStartStop) MessageBox.Show("start MAX");
		
            // we are in the normalize window state
            // get the windows that we minimized previously
            string windows = BFS.ScriptSettings.ReadValue(MinimizedWindowsSetting);
            
            // loop through each setting
            foreach(string window in windows.Split(new char[]{'|'}, StringSplitOptions.RemoveEmptyEntries))
            {
                // try to turn the string into a long value
                // if we can't convert it, go to the next setting
                long windowHandleValue;
                if(!Int64.TryParse(window, out windowHandleValue))
                    continue;
                    
                // restore the window
                BFS.Window.Restore(new IntPtr(windowHandleValue));
            }
            
            // clear the windows that we saved
            BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, string.Empty);
            
            // set the script to MinimizedState
            BFS.ScriptSettings.WriteValue(ScriptStateSetting, MinimizedState);
            if(debugPrintStartStop) MessageBox.Show("finished MAX");
		}
	}
	
	// script is in minimize state if there is no setting, or if the setting is equal to MinimizedState
	private static bool IsScriptInMinimizeState()
	{
		//return true;
		string setting = BFS.ScriptSettings.ReadValue(ScriptStateSetting);
		return (setting.Length == 0) || (setting.Equals(MinimizedState, StringComparison.Ordinal));
	}
}