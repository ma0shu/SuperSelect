using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;

namespace SuperSelect.App.Services;

internal static class AppLogger
{
    private static readonly object SyncRoot = new();
    private static readonly ConcurrentDictionary<string, DateTime> LastExceptionLogByKey =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly ConcurrentQueue<(DateTime Timestamp, string Level, string Message)> PendingEntries = new();
    private static readonly AutoResetEvent FlushSignal = new(false);
    private static readonly CancellationTokenSource FlushCts = new();
    private static readonly Task FlushWorker = Task.Run(FlushLoop);

    private static readonly string LogDirectory = InitializeLogDirectory();
    private static int _throttleCleanupCounter;

    static AppLogger()
    {
        AppDomain.CurrentDomain.ProcessExit += (_, _) => StopFlushWorker();
    }

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
            PendingEntries.Enqueue((DateTime.Now, level, message));
            FlushSignal.Set();
        }
        catch
        {
            // Logging failures must never crash the app.
        }
    }

    private static void FlushLoop()
    {
        while (true)
        {
            _ = FlushSignal.WaitOne();
            FlushPending();

            if (FlushCts.IsCancellationRequested)
            {
                break;
            }
        }

        FlushPending();
    }

    private static void FlushPending()
    {
        if (PendingEntries.IsEmpty)
        {
            return;
        }

        var grouped = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        while (PendingEntries.TryDequeue(out var entry))
        {
            var logPath = GetLogFilePath(entry.Timestamp);
            var line = $"{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{entry.Level}] [PID:{Environment.ProcessId}] {entry.Message}";

            if (!grouped.TryGetValue(logPath, out var lines))
            {
                lines = [];
                grouped[logPath] = lines;
            }

            lines.Add(line);
        }

        if (grouped.Count == 0)
        {
            return;
        }

        try
        {
            lock (SyncRoot)
            {
                Directory.CreateDirectory(LogDirectory);
                foreach (var pair in grouped)
                {
                    File.AppendAllLines(pair.Key, pair.Value, Encoding.UTF8);
                }
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
        MaybeCleanupThrottleMap(now);
        return false;
    }

    private static void MaybeCleanupThrottleMap(DateTime nowUtc)
    {
        if (LastExceptionLogByKey.Count <= 2048)
        {
            return;
        }

        var tick = Interlocked.Increment(ref _throttleCleanupCounter);
        if ((tick & 0xFF) != 0)
        {
            return;
        }

        var expiry = nowUtc - TimeSpan.FromHours(6);
        foreach (var pair in LastExceptionLogByKey)
        {
            if (pair.Value < expiry)
            {
                _ = LastExceptionLogByKey.TryRemove(pair.Key, out _);
            }
        }
    }

    private static void StopFlushWorker()
    {
        try
        {
            FlushCts.Cancel();
            FlushSignal.Set();
            _ = FlushWorker.Wait(1500);
        }
        catch
        {
            // Ignore shutdown races.
        }
        finally
        {
            FlushPending();
        }
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
            // Ignore.
        }

        return root;
    }

    private static string GetLogFilePath(DateTime now)
    {
        return Path.Combine(LogDirectory, $"superselect-{now:yyyyMMdd}.log");
    }
}
