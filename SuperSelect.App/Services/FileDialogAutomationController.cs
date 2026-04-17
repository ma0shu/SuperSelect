using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using SuperSelect.App.Models;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class FileDialogAutomationController
{
    private static readonly string[] FileNameKeywords = ["file name", "filename", "文件名"];
    private static readonly string[] ConfirmKeywords = ["open", "save", "select", "打开", "保存", "选择"];
    private static readonly string[] FileTypeKeywords = ["file type", "save as type", "类型", "文件类型", "保存类型"];
    private static readonly string[] FolderDialogTitleKeywords =
    [
        "select folder",
        "choose folder",
        "browse for folder",
        "打开文件夹",
        "选择文件夹",
        "浏览文件夹",
    ];
    private static readonly string[] FolderConfirmKeywords =
    [
        "select folder",
        "choose folder",
        "打开文件夹",
        "选择文件夹",
    ];
    private static readonly Regex ExtensionRegex = new(@"\*\.[a-zA-Z0-9]+|\.[a-zA-Z0-9]+", RegexOptions.Compiled);
    private static readonly IReadOnlySet<string> EmptyExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private const uint AutomationMessageTimeoutMs = 300;

    private readonly IntPtr _dialogHwnd;
    private readonly object _cacheSyncRoot = new();
    private IntPtr _cachedEditHwnd;
    private IntPtr _cachedConfirmHwnd;
    private IntPtr _cachedFileTypeComboHwnd;
    private string _cachedExtensionFilterText = string.Empty;
    private IReadOnlySet<string> _cachedAllowedExtensions = EmptyExtensions;

    public FileDialogAutomationController(IntPtr dialogHwnd)
    {
        _dialogHwnd = dialogHwnd;
    }

    public static bool IsLikelyFileDialog(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd) || !NativeMethods.IsWindowVisible(hwnd))
        {
            return false;
        }

        var className = NativeMethods.GetWindowClassName(hwnd);
        if (!string.Equals(className, "#32770", StringComparison.Ordinal))
        {
            return false;
        }

        // Modern File Dialogs
        if (NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "DUIViewWndClassName", null) != IntPtr.Zero)
        {
            return true;
        }

        // Older File Dialogs
        if (NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "ComboBoxEx32", null) != IntPtr.Zero ||
            NativeMethods.GetDlgItem(hwnd, 1148) != IntPtr.Zero)
        {
            return true;
        }

        return false;
    }

    public bool TryPrimeSelection(FileCandidate candidate)
    {
        return TryPrimeSelectionAsync(candidate).GetAwaiter().GetResult();
    }

    public async Task<bool> TryPrimeSelectionAsync(FileCandidate candidate, CancellationToken cancellationToken = default)
    {
        if (candidate.IsDirectory)
        {
            return TryNavigateToDirectory(candidate.FullPath);
        }

        var directory = Path.GetDirectoryName(candidate.FullPath);
        var fileName = Path.GetFileName(candidate.FullPath);

        if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        var root = TryGetRoot();
        if (root is null)
        {
            return false;
        }

        var edit = GetFileNameEdit(root);
        var confirm = GetConfirmButton(root);
        if (edit is null || confirm is null)
        {
            InvalidateControlCaches();
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, directory))
        {
            return false;
        }

        if (!TryInvoke(confirm))
        {
            return false;
        }

        try
        {
            await Task.Delay(40, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return false;
        }

        var refreshedRoot = TryGetRoot();
        if (refreshedRoot is not null)
        {
            edit = GetFileNameEdit(refreshedRoot) ?? edit;
        }

        if (!TrySetText(edit, fileName))
        {
            return false;
        }

        try
        {
            await Task.Delay(40, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            return false;
        }

        refreshedRoot = TryGetRoot();
        if (refreshedRoot is not null)
        {
            _ = TrySelectFileInList(refreshedRoot, fileName);
        }

        return true;
    }

    public bool TryConfirmSelection(FileCandidate candidate)
    {
        if (candidate.IsDirectory)
        {
            return TryNavigateToDirectory(candidate.FullPath);
        }

        var root = TryGetRoot();
        if (root is null)
        {
            return false;
        }

        var edit = GetFileNameEdit(root);
        var confirm = GetConfirmButton(root);
        if (edit is null || confirm is null)
        {
            InvalidateControlCaches();
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, candidate.FullPath))
        {
            return false;
        }

        return TryInvoke(confirm);
    }

    public bool TryNavigateToDirectory(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            return false;
        }

        var root = TryGetRoot();
        if (root is null)
        {
            return false;
        }

        var edit = GetFileNameEdit(root);
        var confirm = GetConfirmButton(root);
        if (edit is null || confirm is null)
        {
            InvalidateControlCaches();
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, directoryPath))
        {
            return false;
        }

        return TryInvoke(confirm);
    }

    public IReadOnlySet<string> GetAllowedFileExtensions(string? currentFilterText = null)
    {
        try
        {
            var filterText = currentFilterText;
            if (string.IsNullOrWhiteSpace(filterText))
            {
                var root = TryGetRoot();
                if (root is null)
                {
                    return EmptyExtensions;
                }

                filterText = TryReadFileTypeFilterText(root);
            }

            filterText ??= string.Empty;

            lock (_cacheSyncRoot)
            {
                if (string.Equals(filterText, _cachedExtensionFilterText, StringComparison.OrdinalIgnoreCase))
                {
                    return _cachedAllowedExtensions;
                }
            }

            var parsed = ParseExtensions(filterText);
            lock (_cacheSyncRoot)
            {
                _cachedExtensionFilterText = filterText;
                _cachedAllowedExtensions = parsed;
            }

            return parsed;
        }
        catch
        {
            return EmptyExtensions;
        }
    }

    public string GetCurrentFileTypeFilterText()
    {
        try
        {
            var root = TryGetRoot();
            if (root is null)
            {
                return string.Empty;
            }

            return TryReadFileTypeFilterText(root) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    public bool IsFolderSelectionDialog()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd))
        {
            return false;
        }

        var sb = new System.Text.StringBuilder(256);
        NativeMethods.GetWindowText(_dialogHwnd, sb, sb.Capacity);
        var title = sb.ToString();

        if (!string.IsNullOrWhiteSpace(title) && ContainsAny(title, FolderDialogTitleKeywords))
        {
            return true;
        }

        return false;
    }

    private void TryActivateDialog()
    {
        if (_dialogHwnd != IntPtr.Zero && NativeMethods.IsWindow(_dialogHwnd))
        {
            _ = NativeMethods.SetForegroundWindow(_dialogHwnd);
        }
    }

    private AutomationElement? TryGetRoot()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd))
        {
            InvalidateControlCaches();
            return null;
        }

        try
        {
            return AutomationElement.FromHandle(_dialogHwnd);
        }
        catch
        {
            InvalidateControlCaches();
            return null;
        }
    }

    private AutomationElement? GetFileNameEdit(AutomationElement root)
    {
        IntPtr cachedHandle;
        lock (_cacheSyncRoot)
        {
            cachedHandle = _cachedEditHwnd;
        }

        var cached = TryGetElementFromHandle(cachedHandle);
        if (cached is not null)
        {
            return cached;
        }

        var discovered = FindFileNameEdit(root);
        lock (_cacheSyncRoot)
        {
            _cachedEditHwnd = GetNativeHandle(discovered);
        }

        return discovered;
    }

    private AutomationElement? GetConfirmButton(AutomationElement root)
    {
        IntPtr cachedHandle;
        lock (_cacheSyncRoot)
        {
            cachedHandle = _cachedConfirmHwnd;
        }

        var cached = TryGetElementFromHandle(cachedHandle);
        if (cached is not null)
        {
            return cached;
        }

        var discovered = FindConfirmButton(root);
        lock (_cacheSyncRoot)
        {
            _cachedConfirmHwnd = GetNativeHandle(discovered);
        }

        return discovered;
    }

    private void InvalidateControlCaches()
    {
        lock (_cacheSyncRoot)
        {
            _cachedEditHwnd = IntPtr.Zero;
            _cachedConfirmHwnd = IntPtr.Zero;
            _cachedFileTypeComboHwnd = IntPtr.Zero;
            _cachedExtensionFilterText = string.Empty;
            _cachedAllowedExtensions = EmptyExtensions;
        }
    }

    private static AutomationElement? TryGetElementFromHandle(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd))
        {
            return null;
        }

        try
        {
            return AutomationElement.FromHandle(hwnd);
        }
        catch
        {
            return null;
        }
    }

    private static IntPtr GetNativeHandle(AutomationElement? element)
    {
        if (element is null)
        {
            return IntPtr.Zero;
        }

        try
        {
            return new IntPtr(element.Current.NativeWindowHandle);
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    private static AutomationElement? FindFileNameEdit(AutomationElement root)
    {
        try
        {
            var hwnd = new IntPtr(root.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                var hComboBoxEx32 = NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "ComboBoxEx32", null);
                if (hComboBoxEx32 != IntPtr.Zero)
                {
                    var hComboBox = NativeMethods.FindWindowEx(hComboBoxEx32, IntPtr.Zero, "ComboBox", null);
                    if (hComboBox != IntPtr.Zero)
                    {
                        var hEdit = NativeMethods.FindWindowEx(hComboBox, IntPtr.Zero, "Edit", null);
                        if (hEdit != IntPtr.Zero)
                        {
                            return AutomationElement.FromHandle(hEdit);
                        }
                    }
                }

                var hEditId = NativeMethods.GetDlgItem(hwnd, 1148);
                if (hEditId != IntPtr.Zero)
                {
                    return AutomationElement.FromHandle(hEditId);
                }
            }

            var edits = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

            for (var i = 0; i < edits.Count; i++)
            {
                var edit = edits[i];
                string? name = null;
                try
                {
                    name = edit.Current.Name;
                }
                catch
                {
                    // Ignore per-item transient UIA failure.
                }

                if (!string.IsNullOrWhiteSpace(name) && ContainsAny(name, FileNameKeywords))
                {
                    return edit;
                }
            }

            return edits.Count > 0 ? edits[edits.Count - 1] : null;
        }
        catch
        {
            return null;
        }
    }

    private static AutomationElement? FindConfirmButton(AutomationElement root)
    {
        try
        {
            var hwnd = new IntPtr(root.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                var btn1 = NativeMethods.GetDlgItem(hwnd, 1);
                if (btn1 != IntPtr.Zero)
                {
                    return AutomationElement.FromHandle(btn1);
                }
            }

            var buttons = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));

            for (var i = 0; i < buttons.Count; i++)
            {
                var button = buttons[i];
                string? name = null;
                try
                {
                    name = button.Current.Name;
                }
                catch
                {
                    // Ignore per-item transient UIA failure.
                }

                if (!string.IsNullOrWhiteSpace(name) && ContainsAny(name, ConfirmKeywords))
                {
                    return button;
                }
            }

            return buttons.Count > 0 ? buttons[0] : null;
        }
        catch
        {
            return null;
        }
    }

    private static bool TrySetText(AutomationElement edit, string value)
    {
        try
        {
            var hwnd = new IntPtr(edit.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                if (TrySendTextViaMessage(hwnd, value))
                {
                    return true;
                }
            }

            if (edit.TryGetCurrentPattern(ValuePattern.Pattern, out var patternObj) &&
                patternObj is ValuePattern valuePattern &&
                !valuePattern.Current.IsReadOnly)
            {
                valuePattern.SetValue(value);
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static bool TryInvoke(AutomationElement button)
    {
        try
        {
            var hwnd = new IntPtr(button.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                if (NativeMethods.PostMessage(hwnd, NativeMethods.BM_CLICK, IntPtr.Zero, IntPtr.Zero))
                {
                    return true;
                }
            }

            if (button.TryGetCurrentPattern(InvokePattern.Pattern, out var patternObj) &&
                patternObj is InvokePattern invokePattern)
            {
                invokePattern.Invoke();
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private static bool TrySendTextViaMessage(IntPtr hwnd, string value)
    {
        var result = NativeMethods.SendMessageTimeout(
            hwnd,
            (uint)NativeMethods.WM_SETTEXT,
            IntPtr.Zero,
            value,
            NativeMethods.SMTO_ABORTIFHUNG,
            AutomationMessageTimeoutMs,
            out _);

        return result != IntPtr.Zero;
    }

    private static bool ContainsAny(string value, IReadOnlyList<string> keywords)
    {
        foreach (var keyword in keywords)
        {
            if (value.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasCandidateListControl(AutomationElement root)
    {
        try
        {
            var list = root.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));
            if (list is not null)
            {
                return true;
            }

            var dataGrid = root.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));
            return dataGrid is not null;
        }
        catch
        {
            return false;
        }
    }

    private string? TryReadFileTypeFilterText(AutomationElement root)
    {
        try
        {
            IntPtr cachedComboHandle;
            lock (_cacheSyncRoot)
            {
                cachedComboHandle = _cachedFileTypeComboHwnd;
            }

            var cachedCombo = TryGetElementFromHandle(cachedComboHandle);
            if (cachedCombo is not null)
            {
                var cachedText = GetComboDisplayText(cachedCombo);
                if (!string.IsNullOrWhiteSpace(cachedText))
                {
                    return cachedText;
                }
            }

            var byId = FindFileTypeComboById(root);
            if (byId is not null)
            {
                lock (_cacheSyncRoot)
                {
                    _cachedFileTypeComboHwnd = GetNativeHandle(byId);
                }

                var fromId = GetComboDisplayText(byId);
                if (!string.IsNullOrWhiteSpace(fromId))
                {
                    return fromId;
                }
            }

            var combos = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox));

            for (var i = 0; i < combos.Count; i++)
            {
                var combo = combos[i];
                string? name = null;
                try
                {
                    name = combo.Current.Name;
                }
                catch
                {
                    // Ignore per-item transient UIA failure.
                }

                if (!string.IsNullOrWhiteSpace(name) && ContainsAny(name, FileTypeKeywords))
                {
                    var fromNamedCombo = GetComboDisplayText(combo);
                    if (!string.IsNullOrWhiteSpace(fromNamedCombo))
                    {
                        lock (_cacheSyncRoot)
                        {
                            _cachedFileTypeComboHwnd = GetNativeHandle(combo);
                        }

                        return fromNamedCombo;
                    }
                }
            }

            for (var i = 0; i < combos.Count; i++)
            {
                var combo = combos[i];
                var any = GetComboDisplayText(combo);
                if (!string.IsNullOrWhiteSpace(any) && any.Contains("*.", StringComparison.OrdinalIgnoreCase))
                {
                    lock (_cacheSyncRoot)
                    {
                        _cachedFileTypeComboHwnd = GetNativeHandle(combo);
                    }

                    return any;
                }
            }

            return null;
        }
        catch
        {
            lock (_cacheSyncRoot)
            {
                _cachedFileTypeComboHwnd = IntPtr.Zero;
            }
            return null;
        }
    }

    private static AutomationElement? FindFileTypeComboById(AutomationElement root)
    {
        try
        {
            return root.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "1136")));
        }
        catch
        {
            return null;
        }
    }

    private static bool TrySelectFileInList(AutomationElement root, string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        AutomationElementCollection items;
        try
        {
            items = root.FindAll(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
        }
        catch
        {
            return false;
        }

        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            string? name = null;
            try
            {
                name = item.Current.Name?.Trim();
            }
            catch
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(name) ||
                !name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                if (item.TryGetCurrentPattern(ScrollItemPattern.Pattern, out var scrollObj) &&
                    scrollObj is ScrollItemPattern scrollPattern)
                {
                    scrollPattern.ScrollIntoView();
                }

                if (item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionObj) &&
                    selectionObj is SelectionItemPattern selectionPattern)
                {
                    selectionPattern.Select();
                    return true;
                }

                if (item.TryGetCurrentPattern(InvokePattern.Pattern, out var invokeObj) &&
                    invokeObj is InvokePattern invokePattern)
                {
                    invokePattern.Invoke();
                    return true;
                }
            }
            catch
            {
                continue;
            }
        }

        return false;
    }

    private static string? GetComboDisplayText(AutomationElement combo)
    {
        try
        {
            if (combo.TryGetCurrentPattern(SelectionPattern.Pattern, out var selectionObj) &&
                selectionObj is SelectionPattern selectionPattern)
            {
                var selected = selectionPattern.Current.GetSelection();
                if (selected.Length > 0)
                {
                    var selectedName = selected[0].Current.Name;
                    if (!string.IsNullOrWhiteSpace(selectedName))
                    {
                        return selectedName;
                    }
                }
            }
        }
        catch
        {
            // Ignore and continue.
        }

        try
        {
            var textElement = combo.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));

            if (textElement is not null && !string.IsNullOrWhiteSpace(textElement.Current.Name))
            {
                return textElement.Current.Name;
            }
        }
        catch
        {
            // Ignore and continue.
        }

        try
        {
            return combo.Current.Name;
        }
        catch
        {
            return null;
        }
    }

    private static IReadOnlySet<string> ParseExtensions(string? filterText)
    {
        if (string.IsNullOrWhiteSpace(filterText))
        {
            return EmptyExtensions;
        }

        if (filterText.Contains("*.*", StringComparison.OrdinalIgnoreCase) ||
            filterText.Contains("all files", StringComparison.OrdinalIgnoreCase) ||
            filterText.Contains("所有文件", StringComparison.OrdinalIgnoreCase))
        {
            return EmptyExtensions;
        }

        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var matches = ExtensionRegex.Matches(filterText);
        foreach (Match match in matches)
        {
            var token = match.Value.Trim();
            if (token.StartsWith("*", StringComparison.Ordinal))
            {
                token = token[1..];
            }

            if (!token.StartsWith(".", StringComparison.Ordinal))
            {
                continue;
            }

            token = token.TrimEnd(')', ';', ',', ' ');
            if (token.Length <= 1)
            {
                continue;
            }

            var extBody = token[1..];
            if (extBody.All(char.IsLetterOrDigit))
            {
                result.Add(token.ToLowerInvariant());
            }
        }

        return result.Count == 0 ? EmptyExtensions : result;
    }
}
