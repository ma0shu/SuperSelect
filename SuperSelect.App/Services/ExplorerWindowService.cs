using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using SuperSelect.App.Models;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class ExplorerWindowService
{
    public IReadOnlyList<FileCandidate> GetOpenLocations(string keyword)
    {
        var filter = keyword.Trim();
        var result = new List<FileCandidate>();

        result.AddRange(GetLocationsFromShellWindows(filter));
        if (result.Count == 0)
        {
            result.AddRange(GetLocationsFromWindowAutomation(filter));
        }

        return result
            .GroupBy(item => item.FullPath, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(item => item.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IEnumerable<FileCandidate> GetLocationsFromShellWindows(string filter)
    {
        var result = new List<FileCandidate>();

        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType is null)
        {
            return result;
        }

        object? shellObject = null;
        object? windowsObject = null;

        try
        {
            shellObject = Activator.CreateInstance(shellType);
            if (shellObject is null)
            {
                return result;
            }

            windowsObject = shellType.InvokeMember("Windows", System.Reflection.BindingFlags.InvokeMethod, null, shellObject, null);
            if (windowsObject is null) return result;

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

                    var title = Convert.ToString(windowType.InvokeMember("LocationName", System.Reflection.BindingFlags.GetProperty, null, window, null));
                    result.Add(CreateCandidate(folderPath, title));
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

        return result;
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
