using System;
using System.Windows.Forms;

using System.Text;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;

using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

public static class Log
{
    private const bool singleLogMode = true;
    private const int BUFFER_FLUSH_SIZE = 100;
    private const int BUFFER_FLUSH_MS = 2000;
    private const int BUFFER_CAPACITY = 200;

    private static readonly BlockingCollection<string> _logQueue = new BlockingCollection<string>(BUFFER_CAPACITY);
    private static readonly CancellationTokenSource _cts = new CancellationTokenSource();
    private static Task _flushTask;
    private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _propertyCache = new();
    private static readonly object _lock = new object();
    private static readonly object _emergencyLock = new object();
    private static readonly object _configLock = new object();
    private static DateTime _lastOverflowWarning = DateTime.MinValue;
    private static string _propIndentStr = "";
    private static int COLLECTION_BREAK_CHAR_LIMIT = 105;
    private static readonly string LogFilePath;
    private const string LOG_DIRECTORY = @"C:\Users\mikolaj\Documents\MinimizerLogs";
    private const string FILENAME_PREFIX = "Log_";

    public enum LogLevel
    {
        Error,
        Warning,
        Info,
        Debug
    }

    public static bool EnableMessageBoxes { get; set; } = true;

    public static LogLevel MinimumLogLevel { get; set; } = LogLevel.Debug;

    private static Dictionary<LogLevel, bool> _messageBoxDefaults = new Dictionary<LogLevel, bool>
        {
            { LogLevel.Error, true },
            { LogLevel.Warning, true },
            { LogLevel.Info, false },
            { LogLevel.Debug, false }
        };

    private class NullDisposable : IDisposable
    {
        public static readonly NullDisposable Instance = new();
        public void Dispose() { }
    }

    // TimedOperation factory method
    // usage: using ( Log.T("text") ) { code(); }
    public static IDisposable T(string operationName,
                                LogLevel logLevel = LogLevel.Debug,
                                bool showMessageBox = false,
                                bool? startLog = null,
                                [CallerMemberName] string memberName = "",
                                [CallerLineNumber] int lineNumber = 0)
    {
        return (int)logLevel <= (int)MinimumLogLevel
            ? new TimedOperation(operationName, logLevel, showMessageBox, startLog, memberName, lineNumber)
            : NullDisposable.Instance;
    }

    private class TimedOperation : IDisposable
    {
        private readonly string _operationName;
        private readonly string _memberName;
        private readonly int _lineNumber;
        private readonly LogLevel _logLevel;
        private readonly Stopwatch _sw;
        private readonly bool _showMessageBox;
        private readonly bool _isEnabled;

        public TimedOperation(string operationName,
                            LogLevel logLevel = LogLevel.Debug,
                            bool showMessageBox = false,
                            bool? startLog = null,
                            [CallerMemberName] string memberName = "",
                            [CallerLineNumber] int lineNumber = 0)
        {
            _isEnabled = (int)logLevel <= (int)Log.MinimumLogLevel;
            if (!_isEnabled) return;

            _operationName = operationName;
            _logLevel = logLevel;
            _showMessageBox = showMessageBox;
            _memberName = memberName;
            _lineNumber = lineNumber;

            if (startLog ?? DisplayFusionFunction.logStartTimerDefault)
                UseLogger($"START '{_operationName}'");

            _sw = Stopwatch.StartNew();
        }

        public void Dispose()
        {
            if (!_isEnabled) return;

            _sw.Stop();

            var elapsed = _sw.Elapsed;
            var formattedTime = elapsed.TotalMilliseconds < 1000
                ? $"{elapsed.TotalMilliseconds:0.0000} ms"
                : elapsed.TotalMilliseconds < 60000
                ? elapsed.ToString(@"ss\.ff") + " s"
                : elapsed.ToString(@"hh\:mm\:ss\.f");

            UseLogger($"TIME of '{_operationName}': {formattedTime}");
        }

        private void UseLogger(string message)
        {
            Log.LogInternal(
                _logLevel,
                message,
                variablesFactory: null,
                showMessageBox: _showMessageBox,
                skipHeader: false,
                condition: () => true,
                memberName: _memberName,
                lineNumber: _lineNumber
            );
        }
    } // TimedOperation

    static Log()
    {
        try
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string logDir = singleLogMode ? LOG_DIRECTORY : Path.Combine(LOG_DIRECTORY, date);
            Directory.CreateDirectory(logDir);
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string filename = singleLogMode ? "log.txt" : $"{FILENAME_PREFIX}{timestamp}.txt";
            LogFilePath = Path.Combine(logDir, filename);

            File.WriteAllText(LogFilePath, $"Application Log - {DateTime.Now}");

            StartFlushThread();
            AppDomain.CurrentDomain.ProcessExit += (s, e) => Shutdown();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to initialize logger: {ex.Message}",
                            "Logger Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
        }
    }

    private static void StartFlushThread()
    {
        _flushTask = Task.Run(() =>
        {
            var buffer = new StringBuilder();
            var lastFlush = DateTime.UtcNow;

            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    // get log from queue with timeout
                    if (_logQueue.TryTake(out var entry, BUFFER_FLUSH_MS, _cts.Token))
                    {
                        buffer.Append(entry);
                    }

                    // conditions to write to file
                    bool shouldFlush = buffer.Length > 0 && (
                        buffer.Length >= BUFFER_FLUSH_SIZE * 100 ||  // approx size
                        (DateTime.UtcNow - lastFlush).TotalMilliseconds >= BUFFER_FLUSH_MS
                    );

                    if (shouldFlush)
                    {
                        FlushBuffer(buffer);
                        lastFlush = DateTime.UtcNow;
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }

            // final flush of remaining logs
            FlushBuffer(buffer);
        });
    }

    private static void FlushBuffer(StringBuilder buffer)
    {
        if (buffer.Length == 0) return;

        lock (_lock)
        {
            File.AppendAllText(LogFilePath, buffer.ToString());
            buffer.Clear();
        }
    }

    private static void Shutdown()
    {
        _cts.Cancel();
        _flushTask?.Wait(3000);  // max 3 sec for finalization
    }

    public static void SetMessageBoxDefault(LogLevel level, bool showByDefault)
    {
        lock (_configLock)
        {
            _messageBoxDefaults[level] = showByDefault;
        }
    }

    public static void E(string message,
                        object variables = null,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
           => LogInternal(LogLevel.Error, message, () => variables, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void E(string message,
                        Func<object> variablesFactory,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
           => LogInternal(LogLevel.Error, message, variablesFactory, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void W(string message,
                        object variables = null,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Warning, message, () => variables, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void W(string message,
                        Func<object> variablesFactory,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Warning, message, variablesFactory, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void I(string message,
                        object variables = null,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Info, message, () => variables, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void I(string message,
                        Func<object> variablesFactory,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Info, message, variablesFactory, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void D(string message,
                        object variables = null,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Debug, message, () => variables, showMessageBox, skipHeader, condition, memberName, lineNumber);

    public static void D(string message,
                        Func<object> variablesFactory,
                        bool? showMessageBox = null,
                        bool skipHeader = false,
                        Func<bool> condition = null,
                        [CallerMemberName] string memberName = "",
                        [CallerLineNumber] int lineNumber = 0)
        => LogInternal(LogLevel.Debug, message, variablesFactory, showMessageBox, skipHeader, condition, memberName, lineNumber);

    private static void LogInternal(LogLevel level,
                                   string message,
                                   Func<object> variablesFactory,
                                   bool? showMessageBox,
                                   bool skipHeader,
                                   Func<bool> condition,
                                   string memberName,
                                   int lineNumber)
    {
        // unless condition func was provided assume log should be written
        condition ??= () => true;

        if (condition())
        {
            if ((int)level > (int)MinimumLogLevel) return;

            object variables = null;
            if (variablesFactory != null)
            {
                variables = variablesFactory(); // evaluate variables here
            }

            string formattedMessage = $"{message}";
            if (!skipHeader)
            {
                formattedMessage = $"{memberName}():{lineNumber} {message}";
            }
            string fullMessage = BuildMessageWithVariables(formattedMessage, variables);
            WriteLog(level, fullMessage, showMessageBox, skipHeader);
        }
    }

    private static void WriteLog(LogLevel level, string message, bool? showMessageBox, bool skipHeader)
    {
        try
        {
            string logEntry = skipHeader ?
                $"{message}\n" :
                $"{DateTime.Now:HH:mm:ss} [{level.ToShortString()}]\t{message}\n";

            if (!_logQueue.TryAdd(logEntry, 50)) // timeout 50ms
            {
                HandleQueueOverflow();
            }

            if (ShouldShowMessageBox(level, showMessageBox))
            {
                ShowMessage(level, message, skipHeader);
            }
        }
        catch (Exception ex)
        {
            SafeEmergencyLog($"Critical logging failure: {ex.Message}");
        }
    }

    private static void HandleQueueOverflow()
    {
        try
        {
            // limit frequency of log warnings to max 1 every 5 sec
            if ((DateTime.Now - _lastOverflowWarning).TotalSeconds < 5) return;

            // save info about overflow in safe emergency log
            var errorMsg = "Log queue overflow! Some logs may be lost.";
            SafeEmergencyLog(errorMsg);

            // drain part of main queue
            DrainQueuePartially();

            _lastOverflowWarning = DateTime.Now;
        }
        catch
        {
            // final emergency fallback log
            try { File.AppendAllText("emergency_fallback.log", $"{DateTime.Now:o} UNABLE TO LOG ERRORS\n"); }
            catch { /* ignore */ }
        }
    }

    private static void DrainQueuePartially()
    {
        lock (_emergencyLock)
        {
            // remove 25% oldest logs
            int itemsToRemove = _logQueue.Count / 4;
            for (int i = 0; i < itemsToRemove; i++)
            {
                _logQueue.TryTake(out _);
            }
        }
    }

    private static void SafeEmergencyLog(string message)
    {
        lock (_emergencyLock)
        {
            try
            {
                var tempPath = Path.Combine(LOG_DIRECTORY, "logs_temp.txt");
                File.AppendAllText(tempPath, $"{DateTime.Now:HH:mm:ss} [OVERFLOW_ERR]\t{message}\n");

                // rotate temp log file (>10MB)
                var fi = new FileInfo(tempPath);
                if (fi.Length > 10_000_000)
                {
                    File.Move(tempPath, Path.Combine(LOG_DIRECTORY, $"logs_temp_old_{DateTime.Now:HHmmss}.txt"));
                }
            }
            catch (Exception ex)
            {
                // Final attempt to save logs in non standard location
                try { File.AppendAllText(@"C:\Windows\Temp\emergency.log", $"{DateTime.Now:o}|{ex.Message}\n"); }
                catch { /* Final ignore */ }
            }
        }
    }
    private static bool ShouldShowMessageBox(LogLevel level, bool? showMessageBox)
    {
        lock (_configLock)
        {
            return EnableMessageBoxes && (showMessageBox ?? _messageBoxDefaults[level]);
        }
    }

    private static string BuildMessageWithVariables(string message, object variables)
    {
        var sb = new StringBuilder(message ?? "");

        if (variables != null)
        {
            string variablesString = SerializeVariables(variables);
            if (!string.IsNullOrEmpty(variablesString))
            {
                if (message.Length == 0 && variablesString.Length >= 4) variablesString = variablesString.Substring(4);
                if (sb.Length > 0) sb.AppendLine();
                sb.Append(variablesString);
            }
        }

        return sb.Length > 0 ? sb.ToString() : "Empty log message";
    }

    private static string SerializeVariables(object variables)
    {
        if (variables == null) return string.Empty;

        var sb = new StringBuilder("");
        var type = variables.GetType();
        var properties = _propertyCache.GetOrAdd(type, t => t.GetProperties());

        foreach (var prop in properties)
        {
            var value = prop.GetValue(variables);
            _propIndentStr = new string(' ', prop.Name.Length + 3); // increase indent for ': [' or ': {'
            sb.AppendLine(AddLeadingTabs($"{prop.Name}: {SerializeObject(value)}"));
        }

        if (sb.Length > 1) sb.Length -= 2; // remove newline

        return sb.ToString();
    }

    private static string SerializeObject(object obj, int depth = 0)
    {
        if (depth > 2) return "[...]";
        if (obj == null) return "null";

        return obj switch
        {
            string s => $"\"{s}\"",
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss"),
            IDictionary dict => SerializeDictionary(dict, depth),
            IEnumerable enumerable => SerializeCollection(enumerable, depth),
            IntPtr ptr => $"0x{ptr.ToString()}",
            _ => $"{obj.ToString()}"
        };
    }

    private static void AppendFormattedEntry(StringBuilder sb, string logentry, int count, ref int lineLen)
    {
        if (((lineLen + logentry.Length) > COLLECTION_BREAK_CHAR_LIMIT) && count > 0)
        {
            sb.AppendLine();
            sb.Append(_propIndentStr);
            lineLen = _propIndentStr.Length + logentry.Length;
        }
        else
        {
            lineLen += logentry.Length;
        }

        sb.Append(logentry);
    }

    private static string SerializeCollection(IEnumerable collection, int depth)
    {
        var sb = new StringBuilder("[");
        int count = 0;
        int lineLen = _propIndentStr.Length;
        foreach (var item in collection)
        {
            string logentry = $"{SerializeObject(item, depth + 1)}, ";
            AppendFormattedEntry(sb, logentry, count, ref lineLen);
            count++;
        }

        if (sb.Length > 1) sb.Length -= 2; // remove last comma and space/newline
        sb.Append($"] ({count})");

        return sb.ToString();
    }

    private static string SerializeDictionary(IDictionary dictionary, int depth)
    {
        var sb = new StringBuilder("{");
        int count = 0;
        int lineLen = _propIndentStr.Length;
        foreach (DictionaryEntry entry in dictionary)
        {
            string logentry = $"{SerializeObject(entry.Key)}: {SerializeObject(entry.Value)}, ";

            AppendFormattedEntry(sb, logentry, count, ref lineLen);
            count++;
        }

        if (sb.Length > 1) sb.Length -= 2; // remove last comma and space/newline
        sb.Append($"}} ({count})");
        return sb.ToString();
    }

    private static void ShowMessage(LogLevel level, string message, bool skipHeader)
    {
        MessageBoxIcon icon = level switch
        {
            LogLevel.Error => MessageBoxIcon.Error,
            LogLevel.Warning => MessageBoxIcon.Warning,
            LogLevel.Info => MessageBoxIcon.Information,
            _ => MessageBoxIcon.None
        };

        string caption = skipHeader
            ? "Application Message"
            : $"{level.ToString().ToUpper()} - {DateTime.Now:HH:mm:ss}";

        MessageBox.Show(RemoveLeadingTabs(message), caption, MessageBoxButtons.OK, icon);
    }

    public static string RemoveLeadingTabs(string message)
    {
        if (message == null)
            return null;

        // The regex pattern ^\t+ matches one or more tabs at the beginning of a line.
        // The RegexOptions.Multiline flag makes the ^ anchor match at the start of each line.
        return Regex.Replace(message, @"^(    )+", " ", RegexOptions.Multiline);
    }

    public static string AddLeadingTabs(string message)
    {
        if (message == null)
            return null;

        // The regex pattern ^ matches the beginning of a line.
        // The RegexOptions.Multiline flag makes the ^ anchor match at the start of each line.
        return Regex.Replace(message, @"^", "                ", RegexOptions.Multiline);
    }
} // Log


public static class LogLevelExtensions
{
    public static string ToShortString(this Log.LogLevel logLevel)
    {
        return logLevel switch
        {
            Log.LogLevel.Error => "ERR",
            Log.LogLevel.Warning => "WAR",
            Log.LogLevel.Info => "INF",
            Log.LogLevel.Debug => "DBG",
            _ => logLevel.ToString()
        };
    }
} // LogLevelExtensions
