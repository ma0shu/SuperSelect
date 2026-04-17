using System.Collections.Concurrent;
using System.IO;
using System.Text;

namespace SuperSelect.App.Services;

internal static class AppLogger
{
    private static readonly object SyncRoot = new();
    private static readonly ConcurrentDictionary<string, DateTime> LastExceptionLogByKey =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly string LogDirectory = InitializeLogDirectory();

    public static string LogDirectoryPath => LogDirectory;

    public static void LogInfo(string message)
    {
        Write("INFO", message);
    }

    public static void LogWarning(string message)
    {
        Write("WARN", message);
    }

    public static void LogException(string context, Exception exception, TimeSpan? throttle = null)
    {
        var interval = throttle ?? TimeSpan.Zero;
        if (interval > TimeSpan.Zero && ShouldThrottle(context, exception, interval))
        {
            return;
        }

        var builder = new StringBuilder();
        builder.AppendLine($"Context: {context}");
        builder.AppendLine($"Type: {exception.GetType().FullName}");
        builder.AppendLine($"HResult: 0x{exception.HResult:X8}");
        builder.AppendLine(exception.ToString());

        Write("ERROR", builder.ToString().TrimEnd());
    }

    private static void Write(string level, string message)
    {
        try
        {
            var logPath = GetLogFilePath(DateTime.Now);
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] [PID:{Environment.ProcessId}] {message}";

            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(logPath, line + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch
        {
            // Logging failures must never crash the app.
        }
    }

    private static bool ShouldThrottle(string context, Exception exception, TimeSpan interval)
    {
        var key = $"{context}|{exception.GetType().FullName}|0x{exception.HResult:X8}|{exception.Message}";

        var now = DateTime.UtcNow;
        var last = LastExceptionLogByKey.GetOrAdd(key, now);
        if (now - last < interval)
        {
            return true;
        }

        LastExceptionLogByKey[key] = now;
        return false;
    }

    private static string InitializeLogDirectory()
    {
        var root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SuperSelect",
            "Logs");

        try
        {
            Directory.CreateDirectory(root);
        }
        catch
        {
            // ignore
        }

        return root;
    }

    private static string GetLogFilePath(DateTime now)
    {
        return Path.Combine(LogDirectory, $"superselect-{now:yyyyMMdd}.log");
    }
}
