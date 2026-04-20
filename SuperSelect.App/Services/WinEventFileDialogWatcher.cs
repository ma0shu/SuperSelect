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
    private IntPtr _selectionHook;
    private IntPtr _nameChangeHook;
    private IntPtr _valueChangeHook;
    private IntPtr _activeDialog;
    private readonly DispatcherTimer _teardownTimer;
    private int _contentChangePending;

    public WinEventFileDialogWatcher(Dispatcher dispatcher)
    {
        _dispatcher = dispatcher;
        _teardownTimer = new DispatcherTimer(DispatcherPriority.Send)
        {
            Interval = TimeSpan.FromMilliseconds(32) // 32ms约等于30FPS，对人眼依然是瞬间，但大幅减少线程唤醒
        };
        _teardownTimer.Tick += OnTeardownTimerTick;
    }

    private void OnTeardownTimerTick(object? sender, EventArgs e)
    {
        if (_activeDialog != IntPtr.Zero)
        {
            if (!NativeMethods.IsWindow(_activeDialog) || !NativeMethods.IsWindowVisible(_activeDialog))
            {
                _teardownTimer.Stop();
                _activeDialog = IntPtr.Zero;
                ActiveDialogChanged?.Invoke(null);
            }
        }
    }

    public event Action<IntPtr?>? ActiveDialogChanged;
    public event Action<IntPtr>? ActiveDialogMoved;
    public event Action<IntPtr>? ActiveDialogContentChanged;

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

        _selectionHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_SELECTION,
            NativeMethods.EVENT_OBJECT_SELECTION,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        _nameChangeHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_NAMECHANGE,
            NativeMethods.EVENT_OBJECT_NAMECHANGE,
            IntPtr.Zero,
            _callback,
            0,
            0,
            NativeMethods.WINEVENT_OUTOFCONTEXT | NativeMethods.WINEVENT_SKIPOWNPROCESS);

        _valueChangeHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_VALUECHANGE,
            NativeMethods.EVENT_OBJECT_VALUECHANGE,
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
            return;
        }

        if (_selectionHook == IntPtr.Zero ||
            _nameChangeHook == IntPtr.Zero ||
            _valueChangeHook == IntPtr.Zero)
        {
            AppLogger.LogWarning("WinEventFileDialogWatcher: some content change hooks failed, fallback to polling snapshot updates.");
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

        if (_selectionHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_selectionHook);
            _selectionHook = IntPtr.Zero;
        }

        if (_nameChangeHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_nameChangeHook);
            _nameChangeHook = IntPtr.Zero;
        }

        if (_valueChangeHook != IntPtr.Zero)
        {
            _ = NativeMethods.UnhookWinEvent(_valueChangeHook);
            _valueChangeHook = IntPtr.Zero;
        }

        _teardownTimer.Stop();
        _activeDialog = IntPtr.Zero;
        Interlocked.Exchange(ref _contentChangePending, 0);
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
        if (eventType is NativeMethods.EVENT_OBJECT_SELECTION or
            NativeMethods.EVENT_OBJECT_NAMECHANGE or
            NativeMethods.EVENT_OBJECT_VALUECHANGE)
        {
            if (IsActiveDialogRelatedWindow(hwnd))
            {
                QueueActiveDialogContentChanged();
            }

            return;
        }

        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0)
        {
            return;
        }

        if (eventType == NativeMethods.EVENT_OBJECT_LOCATIONCHANGE)
        {
            if (_activeDialog != IntPtr.Zero && hwnd == _activeDialog)
            {
                _ = _dispatcher.BeginInvoke(
                    DispatcherPriority.Send,
                    new Action(() =>
                    {
                        if (_activeDialog != IntPtr.Zero &&
                            hwnd == _activeDialog &&
                            NativeMethods.IsWindow(_activeDialog))
                        {
                            ActiveDialogMoved?.Invoke(_activeDialog);
                        }
                    }));
            }
            return;
        }

        var priority = eventType == NativeMethods.EVENT_SYSTEM_DIALOGEND
            ? DispatcherPriority.Send
            : DispatcherPriority.Input;

        _ = _dispatcher.BeginInvoke(
            priority,
            new Action(() => EvaluateActiveDialog(eventType, hwnd)));
    }

    private bool IsActiveDialogRelatedWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || _activeDialog == IntPtr.Zero)
        {
            return false;
        }

        if (hwnd == _activeDialog)
        {
            return true;
        }

        var root = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
        return root != IntPtr.Zero && root == _activeDialog;
    }

    private void QueueActiveDialogContentChanged()
    {
        if (Interlocked.Exchange(ref _contentChangePending, 1) != 0)
        {
            return;
        }

        _ = _dispatcher.BeginInvoke(
            DispatcherPriority.Send,
            new Action(() =>
            {
                Interlocked.Exchange(ref _contentChangePending, 0);
                if (_activeDialog == IntPtr.Zero || !NativeMethods.IsWindow(_activeDialog))
                {
                    return;
                }

                ActiveDialogContentChanged?.Invoke(_activeDialog);
            }));
    }

    private void EvaluateActiveDialog(uint eventType, IntPtr sourceHwnd)
    {
        if (_activeDialog != IntPtr.Zero && !NativeMethods.IsWindow(_activeDialog))
        {
            _teardownTimer.Stop();
            _activeDialog = IntPtr.Zero;
            ActiveDialogChanged?.Invoke(null);
        }

        if (_activeDialog != IntPtr.Zero && !NativeMethods.IsWindowVisible(_activeDialog))
        {
            _teardownTimer.Stop();
            _activeDialog = IntPtr.Zero;
            ActiveDialogChanged?.Invoke(null);
        }

        if (eventType == NativeMethods.EVENT_SYSTEM_DIALOGEND &&
            _activeDialog != IntPtr.Zero &&
            sourceHwnd == _activeDialog)
        {
            _teardownTimer.Stop();
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
                _teardownTimer.Stop();
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
        if (candidate != IntPtr.Zero && !_teardownTimer.IsEnabled)
        {
            _teardownTimer.Start();
        }
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
