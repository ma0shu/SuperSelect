using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using SuperSelect.App.Models;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class FileDialogAutomationController
{
    private static readonly string[] FileNameKeywords = ["file name", "filename", "文件名"];
    private static readonly string[] ConfirmKeywords = ["open", "save", "select", "ok", "确定", "打开", "保存", "选择"];
    private static readonly string[] FileTypeKeywords = ["file type", "save as type", "类型", "文件类型", "保存类型"];
    private static readonly string[] FolderDialogTitleKeywords =
    [
        "select folder",
        "choose folder",
        "browse for folder",
        "打开文件夹",
        "选择文件夹",
        "浏览文件夹"
    ];
    private static readonly Regex ExtensionRegex = new(@"\*\.[a-zA-Z0-9]+|\.[a-zA-Z0-9]+", RegexOptions.Compiled);
    private static readonly IReadOnlySet<string> EmptyExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private const uint AutomationMessageTimeoutMs = 300;
    private const uint BffmSetSelectionW = 0x0467;

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
            return false;

        var className = NativeMethods.GetWindowClassName(hwnd);
        if (!string.Equals(className, "#32770", StringComparison.Ordinal))
            return false;

        if (NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "DUIViewWndClassName", null) != IntPtr.Zero)
            return true;

        if (NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "ComboBoxEx32", null) != IntPtr.Zero ||
            NativeMethods.GetDlgItem(hwnd, 1148) != IntPtr.Zero)
            return true;

        var titleBuilder = new System.Text.StringBuilder(256);
        NativeMethods.GetWindowText(hwnd, titleBuilder, titleBuilder.Capacity);
        var title = titleBuilder.ToString();

        if (!string.IsNullOrWhiteSpace(title) && ContainsAny(title, FolderDialogTitleKeywords))
            return true;

        if (NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "SysTreeView32", null) != IntPtr.Zero &&
            NativeMethods.GetDlgItem(hwnd, 1) != IntPtr.Zero)
            return true;

        return false;
    }

    public bool TryPrimeSelection(FileCandidate candidate)
    {
        bool success = TryPrimeSelectionAsync(candidate).GetAwaiter().GetResult();
        if (!success) DumpDialogStateForDebugging("TryPrimeSelection_Failed");
        return success;
    }

    private void DumpDialogStateForDebugging(string tag)
    {
        try
        {
            var dumpSb = new System.Text.StringBuilder();
            dumpSb.AppendLine($"--- Dialog Dump ({tag}) ---");
            dumpSb.AppendLine($"Main Window HWND: {_dialogHwnd}");
            if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd))
            {
                dumpSb.AppendLine("Window is invalid or closed.");
                AppLogger.LogWarning(dumpSb.ToString());
                return;
            }

            var sb = new System.Text.StringBuilder(256);
            NativeMethods.GetWindowText(_dialogHwnd, sb, sb.Capacity);
            dumpSb.AppendLine($"Title: {sb}");

            NativeMethods.GetClassName(_dialogHwnd, sb, sb.Capacity);
            dumpSb.AppendLine($"Class Name: {sb}");

            dumpSb.AppendLine($"IsFolderSelectionDialog(): {IsFolderSelectionDialog()}");

            dumpSb.AppendLine("--- Child Windows ---");
            NativeMethods.EnumChildWindows(_dialogHwnd, (childHwnd, lParam) =>
            {
                int id = NativeMethods.GetDlgCtrlID(childHwnd);
                var clsSb = new System.Text.StringBuilder(256);
                NativeMethods.GetClassName(childHwnd, clsSb, clsSb.Capacity);
                var txtSb = new System.Text.StringBuilder(256);
                NativeMethods.GetWindowText(childHwnd, txtSb, txtSb.Capacity);
                
                dumpSb.AppendLine($"HWND: {childHwnd} | ID: {id} | Class: {clsSb} | Text: {txtSb}");
                return true;
            }, IntPtr.Zero);
            
            AppLogger.LogWarning(dumpSb.ToString());
        }
        catch { }
    }

    public async Task<bool> TryPrimeSelectionAsync(FileCandidate candidate, CancellationToken cancellationToken = default)
    {
        if (candidate.IsDirectory) return TryPrimeDirectorySelection(candidate.FullPath);

        var directory = Path.GetDirectoryName(candidate.FullPath);
        var fileName = Path.GetFileName(candidate.FullPath);

        if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName))
            return false;

        directory = NormalizeDirectoryPath(directory);

        var edit = GetFileNameEdit();
        var confirm = GetConfirmButton();
        if (edit == IntPtr.Zero || confirm == IntPtr.Zero)
        {
            InvalidateControlCaches();
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, directory)) return false;
        if (!TryInvoke(confirm)) return false;

        var sb = new System.Text.StringBuilder(512);
        for (int i = 0; i < 15; i++)
        {
            try { await Task.Delay(80, cancellationToken).ConfigureAwait(false); } catch { return false; }
            var currentEdit = GetFileNameEdit();
            if (currentEdit == IntPtr.Zero) break;
            
            sb.Clear();
            NativeMethods.GetWindowText(currentEdit, sb, sb.Capacity);
            if (sb.ToString() != directory)
            {
                break;
            }
        }

        edit = GetFileNameEdit();
        if (edit == IntPtr.Zero) return false;

        if (!TrySetText(edit, fileName)) return false;

        try { await Task.Delay(40, cancellationToken).ConfigureAwait(false); } catch { return false; }

        return true;
    }

    public bool TryConfirmSelection(FileCandidate candidate)
    {
        bool success = false;
        if (candidate.IsDirectory)
        {
            success = TryConfirmDirectorySelection(candidate.FullPath);
        }
        else
        {
            var edit = GetFileNameEdit();
            var confirm = GetConfirmButton();
            if (edit == IntPtr.Zero || confirm == IntPtr.Zero)
            {
                InvalidateControlCaches();
            }
            else
            {
                TryActivateDialog();
                if (TrySetText(edit, candidate.FullPath))
                {
                    success = TryInvoke(confirm);
                }
            }
        }

        if (!success) DumpDialogStateForDebugging("TryConfirmSelection_Failed");
        return success;
    }

    private bool TryPrimeDirectorySelection(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath)) return false;

        var isFolderDialog = IsFolderSelectionDialog();
        var normalizedPath = NormalizeDirectoryPath(directoryPath);

        TryActivateDialog();

        if (isFolderDialog && TrySetFolderSelectionViaBrowseMessage(normalizedPath))
            return true;

        var edit = GetFileNameEdit();
        if (edit == IntPtr.Zero)
        {
            InvalidateControlCaches();
            return false;
        }

        if (!TrySetText(edit, normalizedPath)) return false;

        if (isFolderDialog) return true;

        var confirm = GetConfirmButton();
        if (confirm == IntPtr.Zero)
        {
            InvalidateControlCaches();
            return false;
        }

        return TryInvoke(confirm);
    }

    private bool TryConfirmDirectorySelection(string directoryPath)
    {
        if (!TryPrimeDirectorySelection(directoryPath)) return false;

        if (!IsFolderSelectionDialog()) return true;

        if (TryClickPrimaryButtonFast()) return true;

        var confirm = GetConfirmButton();
        if (confirm == IntPtr.Zero)
        {
            InvalidateControlCaches();
            return false;
        }

        return TryInvoke(confirm);
    }

    public bool TryNavigateToDirectory(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath)) return false;

        directoryPath = NormalizeDirectoryPath(directoryPath);

        var edit = GetFileNameEdit();
        var confirm = GetConfirmButton();
        if (edit == IntPtr.Zero || confirm == IntPtr.Zero)
        {
            InvalidateControlCaches();
            return false;
        }

        TryActivateDialog();

        if (!TrySetText(edit, directoryPath)) return false;

        return TryInvoke(confirm);
    }

    public IReadOnlySet<string> GetAllowedFileExtensions(string? currentFilterText = null)
    {
        try
        {
            var filterText = currentFilterText;
            if (string.IsNullOrWhiteSpace(filterText))
            {
                filterText = TryReadFileTypeFilterText();
            }

            filterText ??= string.Empty;

            lock (_cacheSyncRoot)
            {
                if (string.Equals(filterText, _cachedExtensionFilterText, StringComparison.OrdinalIgnoreCase))
                    return _cachedAllowedExtensions;
            }

            var parsed = ParseExtensions(filterText);
            lock (_cacheSyncRoot)
            {
                _cachedExtensionFilterText = filterText;
                _cachedAllowedExtensions = parsed;
            }

            return parsed;
        }
        catch { return EmptyExtensions; }
    }

    public string GetCurrentFileTypeFilterText()
    {
        try { return TryReadFileTypeFilterText() ?? string.Empty; }
        catch { return string.Empty; }
    }

    public bool IsFolderSelectionDialog()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd)) return false;

        var sb = new System.Text.StringBuilder(256);
        NativeMethods.GetWindowText(_dialogHwnd, sb, sb.Capacity);
        var title = sb.ToString();

        if (!string.IsNullOrWhiteSpace(title) && ContainsAny(title, FolderDialogTitleKeywords))
            return true;

        bool hasTree = false;
        bool hasDirectUI = false;
        NativeMethods.EnumChildWindows(_dialogHwnd, (child, lParam) =>
        {
            var classNameSb = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(child, classNameSb, classNameSb.Capacity);
            var cls = classNameSb.ToString();
            if (cls == "SysTreeView32") hasTree = true;
            if (cls == "DirectUIHWND") hasDirectUI = true;
            return true;
        }, IntPtr.Zero);

        return hasTree && !hasDirectUI;
    }

    private void TryActivateDialog()
    {
        if (_dialogHwnd != IntPtr.Zero && NativeMethods.IsWindow(_dialogHwnd))
        {
            _ = NativeMethods.SetForegroundWindow(_dialogHwnd);
        }
    }

    private IntPtr GetFileNameEdit()
    {
        IntPtr cachedHandle;
        lock (_cacheSyncRoot) cachedHandle = _cachedEditHwnd;

        if (cachedHandle != IntPtr.Zero && NativeMethods.IsWindow(cachedHandle))
            return cachedHandle;

        var discovered = FindFileNameEdit();
        lock (_cacheSyncRoot) _cachedEditHwnd = discovered;

        return discovered;
    }

    private IntPtr GetConfirmButton()
    {
        IntPtr cachedHandle;
        lock (_cacheSyncRoot) cachedHandle = _cachedConfirmHwnd;

        if (cachedHandle != IntPtr.Zero && NativeMethods.IsWindow(cachedHandle))
            return cachedHandle;

        var discovered = FindConfirmButton();
        lock (_cacheSyncRoot) _cachedConfirmHwnd = discovered;

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

    private IntPtr FindFileNameEdit()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd)) return IntPtr.Zero;

        IntPtr foundEdit = IntPtr.Zero;
        IntPtr firstEdit = IntPtr.Zero;

        NativeMethods.EnumChildWindows(_dialogHwnd, (childHwnd, lParam) =>
        {
            var sb = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(childHwnd, sb, sb.Capacity);
            if (sb.ToString() == "Edit")
            {
                if (firstEdit == IntPtr.Zero) firstEdit = childHwnd;

                int id = NativeMethods.GetDlgCtrlID(childHwnd);
                if (id == 1001 || id == 1148 || id == 1019 || id == 1012)
                {
                    foundEdit = childHwnd;
                    return false; 
                }
            }
            return true;
        }, IntPtr.Zero);

        if (foundEdit != IntPtr.Zero) return foundEdit;
        if (firstEdit != IntPtr.Zero) return firstEdit;

        return IntPtr.Zero;
    }

    private IntPtr FindConfirmButton()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd)) return IntPtr.Zero;

        IntPtr foundBtn = IntPtr.Zero;
        IntPtr firstBtn = IntPtr.Zero;

        NativeMethods.EnumChildWindows(_dialogHwnd, (childHwnd, lParam) =>
        {
            var sb = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(childHwnd, sb, sb.Capacity);
            if (sb.ToString() == "Button")
            {
                if (firstBtn == IntPtr.Zero) firstBtn = childHwnd;

                int id = NativeMethods.GetDlgCtrlID(childHwnd);
                if (id == 1) // IDOK
                {
                    foundBtn = childHwnd;
                    return false; 
                }
            }
            return true;
        }, IntPtr.Zero);

        if (foundBtn != IntPtr.Zero) return foundBtn;

        var btn1 = NativeMethods.GetDlgItem(_dialogHwnd, 1);
        if (btn1 != IntPtr.Zero) return btn1;

        if (firstBtn != IntPtr.Zero) return firstBtn;

        return IntPtr.Zero;
    }

    private static bool TrySetText(IntPtr hwnd, string value)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd)) return false;

        var result = NativeMethods.SendMessageTimeout(
            hwnd,
            (uint)NativeMethods.WM_SETTEXT,
            IntPtr.Zero,
            value,
            NativeMethods.SMTO_ABORTIFHUNG,
            AutomationMessageTimeoutMs,
            out _);

        if (result != IntPtr.Zero)
        {
            var parent = NativeMethods.GetParent(hwnd);
            if (parent != IntPtr.Zero)
            {
                int id = NativeMethods.GetDlgCtrlID(hwnd);
                IntPtr wParam = new IntPtr((NativeMethods.EN_CHANGE << 16) | (id & 0xFFFF));
                NativeMethods.PostMessage(parent, NativeMethods.WM_COMMAND, wParam, hwnd);
            }
            return true;
        }

        return false;
    }

    private static bool TryInvoke(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !NativeMethods.IsWindow(hwnd)) return false;

        return NativeMethods.PostMessage(hwnd, NativeMethods.BM_CLICK, IntPtr.Zero, IntPtr.Zero);
    }

    private bool TryClickPrimaryButtonFast()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd)) return false;

        var okButton = NativeMethods.GetDlgItem(_dialogHwnd, 1);
        if (okButton == IntPtr.Zero || !NativeMethods.IsWindow(okButton)) return false;

        return NativeMethods.PostMessage(okButton, NativeMethods.BM_CLICK, IntPtr.Zero, IntPtr.Zero);
    }

    private static uint GetPidlSize(IntPtr pidl)
    {
        uint size = 0;
        if (pidl == IntPtr.Zero) return 0;
        int cb = System.Runtime.InteropServices.Marshal.ReadInt16(pidl, (int)size);
        while (cb != 0)
        {
            size += (uint)cb;
            cb = System.Runtime.InteropServices.Marshal.ReadInt16(pidl, (int)size);
        }
        size += 2;
        return size;
    }

    private bool TrySetFolderSelectionViaBrowseMessage(string directoryPath)
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd) || string.IsNullOrWhiteSpace(directoryPath))
            return false;

        var normalized = NormalizeDirectoryPath(directoryPath);
        
        NativeMethods.GetWindowThreadProcessId(_dialogHwnd, out uint processId);
        if (processId == 0) return false;

        IntPtr hProcess = NativeMethods.OpenProcess(
            NativeMethods.PROCESS_VM_OPERATION | NativeMethods.PROCESS_VM_WRITE | NativeMethods.PROCESS_VM_READ, 
            false, 
            processId);

        if (hProcess == IntPtr.Zero) return false;

        IntPtr remoteBuffer = IntPtr.Zero;
        IntPtr remotePidl = IntPtr.Zero;
        try
        {
            if (NativeMethods.SHParseDisplayName(normalized, IntPtr.Zero, out var pidl, 0, out _) == 0 && pidl != IntPtr.Zero)
            {
                try
                {
                    uint pidlSize = GetPidlSize(pidl);
                    if (pidlSize > 0)
                    {
                        remotePidl = NativeMethods.VirtualAllocEx(hProcess, IntPtr.Zero, pidlSize, 
                            NativeMethods.MEM_COMMIT | NativeMethods.MEM_RESERVE, NativeMethods.PAGE_READWRITE);

                        if (remotePidl != IntPtr.Zero)
                        {
                            byte[] pidlBytes = new byte[pidlSize];
                            System.Runtime.InteropServices.Marshal.Copy(pidl, pidlBytes, 0, (int)pidlSize);
                            
                            if (NativeMethods.WriteProcessMemory(hProcess, remotePidl, pidlBytes, pidlSize, out _))
                            {
                                var res1 = NativeMethods.SendMessageTimeout(
                                    _dialogHwnd, BffmSetSelectionW, IntPtr.Zero, remotePidl, 
                                    NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out _);
                                if (res1 != IntPtr.Zero) return true;

                                var res2 = NativeMethods.SendMessageTimeout(
                                    _dialogHwnd, 0x0466 /* BffmSetSelectionA */, IntPtr.Zero, remotePidl, 
                                    NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out _);
                                if (res2 != IntPtr.Zero) return true;
                            }
                        }
                    }
                }
                finally
                {
                    NativeMethods.CoTaskMemFree(pidl);
                }
            }

            byte[] pathBytes = System.Text.Encoding.Unicode.GetBytes(normalized + "\0");
            uint bufferSize = (uint)pathBytes.Length;

            remoteBuffer = NativeMethods.VirtualAllocEx(hProcess, IntPtr.Zero, bufferSize, 
                NativeMethods.MEM_COMMIT | NativeMethods.MEM_RESERVE, NativeMethods.PAGE_READWRITE);

            if (remoteBuffer != IntPtr.Zero)
            {
                if (NativeMethods.WriteProcessMemory(hProcess, remoteBuffer, pathBytes, bufferSize, out _))
                {
                    var result = NativeMethods.SendMessageTimeout(
                        _dialogHwnd, BffmSetSelectionW, new IntPtr(1), remoteBuffer, 
                        NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out _);

                    if (result != IntPtr.Zero) return true;
                }
            }

            return false;
        }
        finally
        {
            if (remoteBuffer != IntPtr.Zero) NativeMethods.VirtualFreeEx(hProcess, remoteBuffer, 0, NativeMethods.MEM_RELEASE);
            if (remotePidl != IntPtr.Zero) NativeMethods.VirtualFreeEx(hProcess, remotePidl, 0, NativeMethods.MEM_RELEASE);
            NativeMethods.CloseHandle(hProcess);
        }
    }

    private static string NormalizeDirectoryPath(string directoryPath)
    {
        try 
        { 
            var path = Path.GetFullPath(directoryPath); 
            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()) && !path.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                path += Path.DirectorySeparatorChar;
            return path;
        }
        catch { return directoryPath; }
    }

    private static bool ContainsAny(string value, IReadOnlyList<string> keywords)
    {
        foreach (var keyword in keywords)
        {
            if (value.Contains(keyword, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private string? TryReadFileTypeFilterText()
    {
        IntPtr comboHandle;
        lock (_cacheSyncRoot) comboHandle = _cachedFileTypeComboHwnd;

        if (comboHandle != IntPtr.Zero && NativeMethods.IsWindow(comboHandle))
        {
            var cachedText = GetComboDisplayText(comboHandle);
            if (!string.IsNullOrWhiteSpace(cachedText)) return cachedText;
        }

        var byId = FindFileTypeComboById();
        if (byId != IntPtr.Zero)
        {
            lock (_cacheSyncRoot) _cachedFileTypeComboHwnd = byId;
            var fromId = GetComboDisplayText(byId);
            if (!string.IsNullOrWhiteSpace(fromId)) return fromId;
        }

        return null;
    }

    private IntPtr FindFileTypeComboById()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.IsWindow(_dialogHwnd)) return IntPtr.Zero;

        IntPtr foundCombo = IntPtr.Zero;

        NativeMethods.EnumChildWindows(_dialogHwnd, (childHwnd, lParam) =>
        {
            var sb = new System.Text.StringBuilder(256);
            NativeMethods.GetClassName(childHwnd, sb, sb.Capacity);
            if (sb.ToString() == "ComboBox")
            {
                int id = NativeMethods.GetDlgCtrlID(childHwnd);
                if (id == 1136 || id == 1089)
                {
                    foundCombo = childHwnd;
                    return false; 
                }
            }
            return true;
        }, IntPtr.Zero);

        return foundCombo;
    }

    private static string? GetComboDisplayText(IntPtr combo)
    {
        var idxParam = NativeMethods.SendMessageTimeout(combo, NativeMethods.CB_GETCURSEL, IntPtr.Zero, IntPtr.Zero, NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out var res);
        long idx = (long)res;
        if (idx >= 0 && idxParam != IntPtr.Zero)
        {
            IntPtr lenParam = NativeMethods.SendMessageTimeout(combo, NativeMethods.CB_GETLBTEXTLEN, new IntPtr((int)idx), IntPtr.Zero, NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out var lenRes);
            int len = (int)lenRes;
            if (len > 0)
            {
                var sb = new System.Text.StringBuilder(len + 1);
                var textParam = NativeMethods.SendMessageTimeout(combo, NativeMethods.CB_GETLBTEXT, new IntPtr((int)idx), sb, NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out _);
                if (textParam != IntPtr.Zero)
                {
                    return sb.ToString();
                }
            }
        }
        
        var txtLenParam = NativeMethods.SendMessageTimeout(combo, NativeMethods.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero, NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out var txtLenRes);
        int txtLen = (int)txtLenRes;
        if (txtLen > 0)
        {
            var sb = new System.Text.StringBuilder(txtLen + 1);
            var textParam = NativeMethods.SendMessageTimeout(combo, NativeMethods.WM_GETTEXT, new IntPtr(txtLen + 1), sb, NativeMethods.SMTO_ABORTIFHUNG, AutomationMessageTimeoutMs, out _);
            if (textParam != IntPtr.Zero && sb.Length > 0)
            {
                return sb.ToString();
            }
        }

        return null;
    }

    private static IReadOnlySet<string> ParseExtensions(string? filterText)
    {
        if (string.IsNullOrWhiteSpace(filterText)) return EmptyExtensions;

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
            if (token.StartsWith("*", StringComparison.Ordinal)) token = token[1..];
            if (!token.StartsWith(".", StringComparison.Ordinal)) continue;

            token = token.TrimEnd(')', ';', ',', ' ');
            if (token.Length <= 1) continue;

            var extBody = token[1..];
            if (extBody.All(char.IsLetterOrDigit)) result.Add(token);
        }

        return result.Count == 0 ? EmptyExtensions : result;
    }
}
