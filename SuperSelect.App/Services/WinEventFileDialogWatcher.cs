using System.Windows.Threading;
using SuperSelect.App.Native;

namespace SuperSelect.App.Services;

internal sealed class WinEventFileDialogWatcher : IDisposable
{
    private readonly Dispatcher _dispatcher;
    private NativeMethods.WinEventDelegate? _callback;
    private IntPtr _foregroundHook;
    private IntPtr _dialogHook;
    private IntPtr _objectShowHook;
    private IntPtr _locationHook;
    private IntPtr _activeDialog;

    public WinEventFileDialogWatcher(Dispatcher dispatcher)
    {
        _dispatcher = dispatcher;
    }

    public event Action<IntPtr?>? ActiveDialogChanged;
    public event Action<IntPtr>? ActiveDialogMoved;

    public bool IsRunning =>
        _foregroundHook != IntPtr.Zero &&
        _dialogHook != IntPtr.Zero &&
        _objectShowHook != IntPtr.Zero &&
        _locationHook != IntPtr.Zero;

    public void Start()
    {
        if (IsRunning)
        {
            return;
        }

        _callback = OnWinEvent;

        _foregroundHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            NativeMethods.EVENT_SYSTEM_FOREGROUND,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        _dialogHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_SYSTEM_DIALOGSTART,
            NativeMethods.EVENT_SYSTEM_DIALOGEND,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        _objectShowHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_SHOW,
            NativeMethods.EVENT_OBJECT_SHOW,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        _locationHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_LOCATIONCHANGE,
            NativeMethods.EVENT_OBJECT_LOCATIONCHANGE,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        if (_foregroundHook == IntPtr.Zero ||
            _dialogHook == IntPtr.Zero ||
            _objectShowHook == IntPtr.Zero ||
            _locationHook == IntPtr.Zero)
        {
            Stop();
        }
    }

    public void Stop()
    {
        if (_foregroundHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_foregroundHook);
            _foregroundHook = IntPtr.Zero;
        }

        if (_dialogHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_dialogHook);
            _dialogHook = IntPtr.Zero;
        }

        if (_objectShowHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_objectShowHook);
            _objectShowHook = IntPtr.Zero;
        }

        if (_locationHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_locationHook);
            _locationHook = IntPtr.Zero;
        }

        _activeDialog = IntPtr.Zero;
    }

    public void Dispose()
    {
        Stop();
        GC.SuppressFinalize(this);
    }

    private void OnWinEvent(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint dwEventThread,
        uint dwmsEventTime)
    {
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0)
        {
            return;
        }

        if (eventType == NativeMethods.EVENT_OBJECT_LOCATIONCHANGE)
        {
            if (_activeDialog != IntPtr.Zero && hwnd == _activeDialog)
            {
                _ = _dispatcher.BeginInvoke(
                    DispatcherPriority.Render,
                    new Action(() => EvaluateActiveDialog(eventType, hwnd)));
            }
            return;
        }

        _ = _dispatcher.BeginInvoke(
            DispatcherPriority.Input,
            new Action(() => EvaluateActiveDialog(eventType, hwnd)));
    }

    private void EvaluateActiveDialog(uint eventType, IntPtr sourceHwnd)
    {
        if (eventType == NativeMethods.EVENT_OBJECT_LOCATIONCHANGE &&
            _activeDialog != IntPtr.Zero &&
            sourceHwnd == _activeDialog &&
            NativeMethods.IsWindow(_activeDialog))
        {
            ActiveDialogMoved?.Invoke(_activeDialog);
            return;
        }

        if (eventType == NativeMethods.EVENT_OBJECT_LOCATIONCHANGE)
        {
            return;
        }

        if (_activeDialog != IntPtr.Zero && !NativeMethods.IsWindow(_activeDialog))
        {
            _activeDialog = IntPtr.Zero;
            ActiveDialogChanged?.Invoke(null);
        }

        if (_activeDialog != IntPtr.Zero && !NativeMethods.IsWindowVisible(_activeDialog))
        {
            _activeDialog = IntPtr.Zero;
            ActiveDialogChanged?.Invoke(null);
        }

        if (eventType == NativeMethods.EVENT_SYSTEM_DIALOGEND &&
            _activeDialog != IntPtr.Zero &&
            sourceHwnd == _activeDialog)
        {
            _activeDialog = IntPtr.Zero;
            ActiveDialogChanged?.Invoke(null);
            return;
        }

        var candidate = ResolveCandidateDialog(eventType, sourceHwnd);

        if (candidate == IntPtr.Zero)
        {
            if (_activeDialog != IntPtr.Zero && NativeMethods.IsWindow(_activeDialog))
            {
                return;
            }

            if (_activeDialog != IntPtr.Zero)
            {
                _activeDialog = IntPtr.Zero;
                ActiveDialogChanged?.Invoke(null);
            }

            return;
        }

        if (candidate == _activeDialog)
        {
            return;
        }

        _activeDialog = candidate;
        ActiveDialogChanged?.Invoke(candidate);
    }

    private static IntPtr ResolveDialog(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        var rootWindow = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
        if (rootWindow == IntPtr.Zero)
        {
            return IntPtr.Zero;
        }

        _ = NativeMethods.GetWindowThreadProcessId(rootWindow, out var pid);
        if (pid == (uint)Environment.ProcessId)
        {
            return IntPtr.Zero;
        }

        return FileDialogAutomationController.IsLikelyFileDialog(rootWindow) ? rootWindow : IntPtr.Zero;
    }

    private static IntPtr ResolveCandidateDialog(uint eventType, IntPtr sourceHwnd)
    {
        if (eventType == NativeMethods.EVENT_SYSTEM_DIALOGSTART ||
            eventType == NativeMethods.EVENT_OBJECT_SHOW)
        {
            var fromSource = ResolveDialog(sourceHwnd);
            if (fromSource != IntPtr.Zero)
            {
                return fromSource;
            }
        }

        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground != IntPtr.Zero)
        {
            var fromForeground = ResolveDialog(foreground);
            if (fromForeground != IntPtr.Zero)
            {
                return fromForeground;
            }
        }

        return ResolveDialog(sourceHwnd);
    }
}
