using System;
using System.Drawing;

// The 'windowHandle' parameter will contain the window handle for the:
//   - Active window when run by hotkey
//   - Window Location target when run by a Window Location rule
//   - TitleBar Button owner when run by a TitleBar Button
//   - Jump List owner when run from a Taskbar Jump List
//   - Currently focused window if none of these match
public static class DisplayFusionFunction
{
	private const string ScriptStateSetting = "CursorMonitorScriptState";
	private const string MinimizedWindowsSetting = "CursorMonitorMinimizedWindows";
	private const string MinimizedState = "0";
	private const string NormalizeState = "1";
	
	public static void Run(IntPtr windowHandle)
	{
		//check to see if we are minimizing 
		if(IsScriptInMinimizeState())
		{
			//this will store the windows that we are minimizing so we can restore them later
			string minimizedWindows = "";
			
			//get the monitor that the cursor is on
			uint monitorId = BFS.Monitor.GetMonitorIDByXY(BFS.Input.GetMousePositionX(), BFS.Input.GetMousePositionY());
			
			//loop through all the visible windows on the cursor monitor
			foreach(IntPtr window in BFS.Window.GetVisibleWindowHandlesByMonitor(monitorId))
			{
				//skip any special DisplayFusion window (taskbar, titlebar buttons)
				//skip special explorer.exe windows (icons, search)
				if(BFS.Window.GetClass(window).StartsWith("DF", StringComparison.Ordinal) ||
					BFS.Window.GetClass(window).Equals("WorkerW", StringComparison.Ordinal) ||
					BFS.Window.GetClass(window).Equals("SearchPane", StringComparison.Ordinal))
				{
						continue;
				}
				
				//minimize the window
				BFS.Window.Minimize(window);
				
				//add the window to the list of windows
				minimizedWindows += window.ToInt64().ToString() + "|";
			}
			
			//save the list of windows we minimized
			BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, minimizedWindows);
			
			//set the script state to NormalizeState
			BFS.ScriptSettings.WriteValue(ScriptStateSetting, NormalizeState);
			
			//exit the script
			return;
		}
		
		//if we got here, we are in the normalize window state
		//get the windows that we minimized previously
		string windows = BFS.ScriptSettings.ReadValue(MinimizedWindowsSetting);
		
		//loop through each setting
		foreach(string window in windows.Split(new char[]{'|'}, StringSplitOptions.RemoveEmptyEntries))
		{
			//try to turn the string into a long value
			//if we can't convert it, go to the next setting
			long windowHandleValue;
			if(!Int64.TryParse(window, out windowHandleValue))
				continue;
				
			//restore the window
			BFS.Window.Restore(new IntPtr(windowHandleValue));
		}
		
		//clear the windows that we saved
		BFS.ScriptSettings.WriteValue(MinimizedWindowsSetting, string.Empty);
		
		//set the script to MinimizedState
		BFS.ScriptSettings.WriteValue(ScriptStateSetting, MinimizedState);
	}
	
	//script is in minimize state if there is no setting, or if the setting is equal to MinimizedState
	private static bool IsScriptInMinimizeState()
	{
		string setting = BFS.ScriptSettings.ReadValue(ScriptStateSetting);
		return (setting.Length == 0) || (setting.Equals(MinimizedState, StringComparison.Ordinal));
	}
}