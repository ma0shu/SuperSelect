using System.IO;
using System.Threading;
using System.Windows.Automation;
using SuperSelect.App.Models;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class FileDialogAutomationController
{
    private static readonly string[] FileNameKeywords = ["file name", "filename", "文件名"];
    private static readonly string[] ConfirmKeywords = ["open", "save", "select", "打开", "保存", "选择"];

    private readonly IntPtr _dialogHwnd;

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

        try
        {
            var root = AutomationElement.FromHandle(hwnd);
            return root is not null
                && FindFileNameEdit(root) is not null
                && FindConfirmButton(root) is not null;
        }
        catch
        {
            return false;
        }
    }

    public bool TryPrimeSelection(FileCandidate candidate)
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

        var edit = FindFileNameEdit(root);
        var confirm = FindConfirmButton(root);
        if (edit is null || confirm is null)
        {
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

        Thread.Sleep(40);
        return TrySetText(edit, fileName);
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

        var edit = FindFileNameEdit(root);
        var confirm = FindConfirmButton(root);
        if (edit is null || confirm is null)
        {
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

        var edit = FindFileNameEdit(root);
        var confirm = FindConfirmButton(root);
        if (edit is null || confirm is null)
        {
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, directoryPath))
        {
            return false;
        }

        return TryInvoke(confirm);
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
            return null;
        }

        try
        {
            return AutomationElement.FromHandle(_dialogHwnd);
        }
        catch
        {
            return null;
        }
    }

    private static AutomationElement? FindFileNameEdit(AutomationElement root)
    {
        var byId = root.FindFirst(
            TreeScope.Descendants,
            new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "1148")));

        if (byId is not null)
        {
            return byId;
        }

        var edits = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

        for (var i = 0; i < edits.Count; i++)
        {
            var edit = edits[i];
            var name = edit.Current.Name;
            if (!string.IsNullOrWhiteSpace(name) && ContainsAny(name, FileNameKeywords))
            {
                return edit;
            }
        }

        return edits.Count > 0 ? edits[edits.Count - 1] : null;
    }

    private static AutomationElement? FindConfirmButton(AutomationElement root)
    {
        var byId = root.FindFirst(
            TreeScope.Descendants,
            new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                new PropertyCondition(AutomationElement.AutomationIdProperty, "1")));

        if (byId is not null)
        {
            return byId;
        }

        var buttons = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));

        for (var i = 0; i < buttons.Count; i++)
        {
            var button = buttons[i];
            var name = button.Current.Name;
            if (!string.IsNullOrWhiteSpace(name) && ContainsAny(name, ConfirmKeywords))
            {
                return button;
            }
        }

        return buttons.Count > 0 ? buttons[0] : null;
    }

    private static bool TrySetText(AutomationElement edit, string value)
    {
        try
        {
            if (edit.TryGetCurrentPattern(ValuePattern.Pattern, out var patternObj) &&
                patternObj is ValuePattern valuePattern &&
                !valuePattern.Current.IsReadOnly)
            {
                valuePattern.SetValue(value);
                return true;
            }

            var hwnd = new IntPtr(edit.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                _ = NativeMethods.SendMessage(hwnd, NativeMethods.WM_SETTEXT, IntPtr.Zero, value);
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
            if (button.TryGetCurrentPattern(InvokePattern.Pattern, out var patternObj) &&
                patternObj is InvokePattern invokePattern)
            {
                invokePattern.Invoke();
                return true;
            }

            var hwnd = new IntPtr(button.Current.NativeWindowHandle);
            if (hwnd != IntPtr.Zero)
            {
                _ = NativeMethods.SendMessage(hwnd, NativeMethods.BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return true;
            }
        }
        catch
        {
            return false;
        }

        return false;
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
}
