using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using SuperSelect.App.Models;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class ExplorerWindowService
{
    private static readonly TimeSpan SnapshotTtl = TimeSpan.FromSeconds(1);
    private static readonly TimeSpan EmptySnapshotNegativeTtl = TimeSpan.FromSeconds(2);

    private readonly object _snapshotSync = new();
    private IReadOnlyList<FileCandidate> _cachedSnapshot = [];
    private DateTime _lastSnapshotUtc = DateTime.MinValue;
    private DateTime _emptySnapshotValidUntilUtc = DateTime.MinValue;
    private int _refreshInFlight;

    public ExplorerWindowService()
    {
        EnsureSnapshotFreshAsync(force: true);
    }

    public IReadOnlyList<FileCandidate> GetOpenLocations(string keyword)
    {
        var filter = keyword.Trim();
        var nowUtc = DateTime.UtcNow;

        IReadOnlyList<FileCandidate>? snapshotToUse = null;
        var skipRefreshForNegativeCache = false;

        lock (_snapshotSync)
        {
            if (nowUtc - _lastSnapshotUtc <= SnapshotTtl)
            {
                snapshotToUse = _cachedSnapshot;
            }
            else if (_cachedSnapshot.Count == 0 && nowUtc <= _emptySnapshotValidUntilUtc)
            {
                snapshotToUse = _cachedSnapshot;
                skipRefreshForNegativeCache = true;
            }
        }

        if (snapshotToUse is not null)
        {
            if (!skipRefreshForNegativeCache)
            {
                EnsureSnapshotFreshAsync(force: false);
            }

            return FilterSnapshot(snapshotToUse, filter);
        }

        var refreshed = RefreshSnapshotNow();
        return FilterSnapshot(refreshed, filter);
    }

    public IReadOnlyList<FileCandidate> GetOpenLocationsCached(string keyword)
    {
        EnsureSnapshotFreshAsync(force: false);

        IReadOnlyList<FileCandidate> snapshot;
        lock (_snapshotSync)
        {
            snapshot = _cachedSnapshot;
        }

        return FilterSnapshot(snapshot, keyword.Trim());
    }

    public bool HasOpenLocationsCached()
    {
        EnsureSnapshotFreshAsync(force: false);

        lock (_snapshotSync)
        {
            return _cachedSnapshot.Count > 0;
        }
    }

    public void EnsureSnapshotFreshAsync(bool force = false)
    {
        var nowUtc = DateTime.UtcNow;
        bool shouldRefresh;

        lock (_snapshotSync)
        {
            shouldRefresh = force || nowUtc - _lastSnapshotUtc > SnapshotTtl;
            if (!force && _cachedSnapshot.Count == 0 && nowUtc <= _emptySnapshotValidUntilUtc)
            {
                shouldRefresh = false;
            }
        }

        if (!shouldRefresh)
        {
            return;
        }

        if (Interlocked.CompareExchange(ref _refreshInFlight, 1, 0) != 0)
        {
            return;
        }

        _ = Task.Run(() =>
        {
            try
            {
                _ = RefreshSnapshotNow();
            }
            catch
            {
                // Keep stale snapshot on refresh errors.
            }
            finally
            {
                Interlocked.Exchange(ref _refreshInFlight, 0);
            }
        });
    }

    private IReadOnlyList<FileCandidate> RefreshSnapshotNow()
    {
        var discovered = DiscoverOpenLocations();
        var nowUtc = DateTime.UtcNow;

        lock (_snapshotSync)
        {
            _cachedSnapshot = discovered;
            _lastSnapshotUtc = nowUtc;
            _emptySnapshotValidUntilUtc = discovered.Count == 0
                ? nowUtc + EmptySnapshotNegativeTtl
                : DateTime.MinValue;

            return _cachedSnapshot;
        }
    }

    private static IReadOnlyList<FileCandidate> DiscoverOpenLocations()
    {
        var result = new List<FileCandidate>();

        result.AddRange(GetLocationsFromShellWindows(string.Empty));
        if (result.Count == 0)
        {
            result.AddRange(GetLocationsFromWindowAutomation(string.Empty));
        }

        return DeduplicatePreservingOrder(result);
    }

    private static IReadOnlyList<FileCandidate> FilterSnapshot(
        IReadOnlyList<FileCandidate> snapshot,
        string filter)
    {
        if (snapshot.Count == 0 || string.IsNullOrWhiteSpace(filter))
        {
            return snapshot;
        }

        return snapshot
            .Where(candidate => MatchesFilter(candidate.FullPath, filter))
            .ToList();
    }

    private static IEnumerable<FileCandidate> GetLocationsFromShellWindows(string filter)
    {
        var zOrderByHandle = GetExplorerWindowZOrderMap();
        var result = new List<(int Rank, FileCandidate Candidate)>();

        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType is null)
        {
            return [];
        }

        object? shellObject = null;
        object? windowsObject = null;

        try
        {
            shellObject = Activator.CreateInstance(shellType);
            if (shellObject is null)
            {
                return [];
            }

            windowsObject = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellObject, null);
            if (windowsObject is null) return [];

            var windowsType = windowsObject.GetType();
            var count = (int)windowsType.InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, windowsObject, null)!;

            for (var i = 0; i < count; i++)
            {
                object? window = null;
                try
                {
                    window = windowsType.InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, windowsObject, new object[] { i });
                    if (window is null) continue;

                    var windowType = window.GetType();
                    var processPath = Convert.ToString(windowType.InvokeMember("FullName", System.Reflection.BindingFlags.GetProperty, null, window, null));

                    if (string.IsNullOrWhiteSpace(processPath) ||
                        !processPath.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var folderPath = ResolveExplorerPathFromShell(window, windowType);
                    if (string.IsNullOrWhiteSpace(folderPath))
                    {
                        continue;
                    }

                    if (!MatchesFilter(folderPath, filter))
                    {
                        continue;
                    }

                    var handle = TryGetShellWindowHandle(window, windowType);
                    var rank = zOrderByHandle.TryGetValue(handle, out var value)
                        ? value
                        : int.MaxValue;
                    var title = Convert.ToString(windowType.InvokeMember("LocationName", System.Reflection.BindingFlags.GetProperty, null, window, null));
                    result.Add((rank, CreateCandidate(folderPath, title)));
                }
                catch
                {
                    // Ignore per-window failure.
                }
                finally
                {
                    if (window is not null && Marshal.IsComObject(window))
                    {
                        _ = Marshal.FinalReleaseComObject(window);
                    }
                }
            }
        }
        catch
        {
            // Ignore COM failure and fallback to UI Automation.
        }
        finally
        {
            if (windowsObject is not null && Marshal.IsComObject(windowsObject))
            {
                _ = Marshal.FinalReleaseComObject(windowsObject);
            }

            if (shellObject is not null && Marshal.IsComObject(shellObject))
            {
                _ = Marshal.FinalReleaseComObject(shellObject);
            }
        }

        return result
            .OrderBy(item => item.Rank)
            .Select(item => item.Candidate)
            .ToList();
    }

    private static IEnumerable<FileCandidate> GetLocationsFromWindowAutomation(string filter)
    {
        var handles = new List<IntPtr>();
        _ = NativeMethods.EnumWindows(
            (hwnd, _) =>
            {
                if (!NativeMethods.IsWindowVisible(hwnd))
                {
                    return true;
                }

                var className = NativeMethods.GetWindowClassName(hwnd);
                if (string.Equals(className, "CabinetWClass", StringComparison.Ordinal) ||
                    string.Equals(className, "ExploreWClass", StringComparison.Ordinal))
                {
                    handles.Add(hwnd);
                }

                return true;
            },
            IntPtr.Zero);

        var result = new List<FileCandidate>();
        foreach (var hwnd in handles)
        {
            var folderPath = ResolveExplorerPathFromAutomation(hwnd);
            if (string.IsNullOrWhiteSpace(folderPath) || !MatchesFilter(folderPath, filter))
            {
                continue;
            }

            result.Add(CreateCandidate(folderPath, Path.GetFileName(folderPath)));
        }

        return result;
    }

    private static FileCandidate CreateCandidate(string path, string? title)
    {
        return new FileCandidate
        {
            FullPath = path,
            DisplayName = string.IsNullOrWhiteSpace(title) ? path : title!,
            SecondaryText = path,
            IsDirectory = true,
            Source = CandidateSource.Explorer,
        };
    }

    private static Dictionary<IntPtr, int> GetExplorerWindowZOrderMap()
    {
        var rank = 0;
        var orderByHandle = new Dictionary<IntPtr, int>();
        _ = NativeMethods.EnumWindows(
            (hwnd, _) =>
            {
                if (!NativeMethods.IsWindowVisible(hwnd))
                {
                    return true;
                }

                var className = NativeMethods.GetWindowClassName(hwnd);
                if (string.Equals(className, "CabinetWClass", StringComparison.Ordinal) ||
                    string.Equals(className, "ExploreWClass", StringComparison.Ordinal))
                {
                    orderByHandle[hwnd] = rank++;
                }

                return true;
            },
            IntPtr.Zero);

        return orderByHandle;
    }

    private static IntPtr TryGetShellWindowHandle(object window, Type windowType)
    {
        try
        {
            var raw = windowType.InvokeMember("HWND", System.Reflection.BindingFlags.GetProperty, null, window, null);
            return raw switch
            {
                int value32 => new IntPtr(value32),
                long value64 => new IntPtr(value64),
                IntPtr ptr => ptr,
                _ => IntPtr.Zero,
            };
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    private static IReadOnlyList<FileCandidate> DeduplicatePreservingOrder(IEnumerable<FileCandidate> items)
    {
        var result = new List<FileCandidate>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in items)
        {
            if (!seen.Add(item.FullPath))
            {
                continue;
            }

            result.Add(item);
        }

        return result;
    }

    private static string? ResolveExplorerPathFromShell(object window, Type windowType)
    {
        string? folderPath = null;

        try
        {
            var document = windowType.InvokeMember("Document", System.Reflection.BindingFlags.GetProperty, null, window, null);
            if (document != null)
            {
                var folder = document.GetType().InvokeMember("Folder", System.Reflection.BindingFlags.GetProperty, null, document, null);
                if (folder != null)
                {
                    var self = folder.GetType().InvokeMember("Self", System.Reflection.BindingFlags.GetProperty, null, folder, null);
                    if (self != null)
                    {
                        folderPath = Convert.ToString(self.GetType().InvokeMember("Path", System.Reflection.BindingFlags.GetProperty, null, self, null));
                    }
                }
            }
        }
        catch
        {
            // Ignore and fallback to URL parsing.
        }

        if (IsExistingDirectory(folderPath))
        {
            return Path.GetFullPath(folderPath!);
        }

        try
        {
            var locationUrl = Convert.ToString(windowType.InvokeMember("LocationURL", System.Reflection.BindingFlags.GetProperty, null, window, null));
            folderPath = TryParseFileUrl(locationUrl);
        }
        catch
        {
            return null;
        }

        return IsExistingDirectory(folderPath) ? Path.GetFullPath(folderPath!) : null;
    }

    private static string? ResolveExplorerPathFromAutomation(IntPtr hwnd)
    {
        AutomationElement? root;
        try
        {
            root = AutomationElement.FromHandle(hwnd);
        }
        catch
        {
            return null;
        }

        if (root is null)
        {
            return null;
        }

        var edits = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

        for (var i = 0; i < edits.Count; i++)
        {
            var edit = edits[i];
            var byValue = TryReadValue(edit);
            var normalized = NormalizeExplorerPathText(byValue);
            if (normalized is not null)
            {
                return normalized;
            }

            var byName = NormalizeExplorerPathText(edit.Current.Name);
            if (byName is not null)
            {
                return byName;
            }
        }

        return null;
    }

    private static string? TryReadValue(AutomationElement element)
    {
        try
        {
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var patternObject) &&
                patternObject is ValuePattern valuePattern)
            {
                return valuePattern.Current.Value;
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static string? NormalizeExplorerPathText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var value = text
            .Replace("\u200e", string.Empty)
            .Replace("\u200f", string.Empty)
            .Trim();

        foreach (var prefix in new[] { "Address:", "Address", "地址:", "地址" })
        {
            if (value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                value = value[prefix.Length..].Trim();
                break;
            }
        }

        if (value.StartsWith("file://", StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(value, UriKind.Absolute, out var uri) &&
            uri.IsFile)
        {
            value = Uri.UnescapeDataString(uri.LocalPath);
        }

        value = value.Replace('/', '\\');
        if (!LooksLikePath(value) || !Directory.Exists(value))
        {
            return null;
        }

        try
        {
            return Path.GetFullPath(value);
        }
        catch
        {
            return null;
        }
    }

    private static bool LooksLikePath(string value)
    {
        if (value.Length < 3)
        {
            return false;
        }

        if (value.StartsWith(@"\\", StringComparison.Ordinal))
        {
            return true;
        }

        return value.Length >= 3
            && char.IsLetter(value[0])
            && value[1] == ':'
            && value[2] == '\\';
    }

    private static bool IsExistingDirectory(string? path)
    {
        return !string.IsNullOrWhiteSpace(path) && Directory.Exists(path);
    }

    private static bool MatchesFilter(string path, string filter)
    {
        return string.IsNullOrWhiteSpace(filter) ||
            path.Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private static string? TryParseFileUrl(string? locationUrl)
    {
        if (string.IsNullOrWhiteSpace(locationUrl))
        {
            return null;
        }

        if (!Uri.TryCreate(locationUrl, UriKind.Absolute, out var uri) || !uri.IsFile)
        {
            return null;
        }

        var localPath = Uri.UnescapeDataString(uri.LocalPath);
        return string.IsNullOrWhiteSpace(localPath) ? null : localPath;
    }
}
