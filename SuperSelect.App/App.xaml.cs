using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Interop;
using System.Windows.Threading;
using SuperSelect.App.Services;
using SuperSelect.App.Views;

namespace SuperSelect.App;

public partial class App : System.Windows.Application
{
    private WinEventFileDialogWatcher? _dialogWatcher;
    private OverlayWindow? _overlayWindow;
    private AppTrayIcon? _trayIcon;
    private HwndSource? _hotkeySource;
    private bool _isExitRequested;
    private bool _isSystemDarkMode;
    private DispatcherTimer? _dialogDetachDebounceTimer;
    private IntPtr _activeDialogHwnd;

    private EverythingService? _everythingService;
    private TrayRepository? _trayRepository;
    private ExplorerWindowService? _explorerWindowService;
    private UserPreferencesRepository? _preferencesRepository;
    private DateTime _lastMemoryTrimUtc = DateTime.MinValue;
    private static readonly TimeSpan MemoryTrimCooldown = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan MemoryTrimStartDelay = TimeSpan.FromMilliseconds(1000);
    private const long MemoryTrimManagedThresholdBytes = 64L * 1024 * 1024;
    private const string DialogDetachedTrimReason = "DialogDetached";
    private static readonly TimeSpan DialogDetachDebounceInterval = TimeSpan.FromMilliseconds(90);
    private const int MainWindowHotkeyId = 0x5353;
    private const uint MainWindowHotkeyModifiers = Native.NativeMethods.MOD_CONTROL | Native.NativeMethods.MOD_SHIFT;
    private const uint MainWindowHotkeyVKey = (uint)'M';
    private const string MainWindowHotkeyDisplayName = "Ctrl+Shift+M";

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnTaskSchedulerUnobservedTaskException;

        _everythingService = new EverythingService();
        _trayRepository = new TrayRepository();
        _explorerWindowService = new ExplorerWindowService();
        _preferencesRepository = new UserPreferencesRepository();

        AppLogger.LogInfo($"Application startup. .NET={Environment.Version}; OS={Environment.OSVersion}");
        AppLogger.LogInfo($"Log directory: {AppLogger.LogDirectoryPath}");

        _overlayWindow = new OverlayWindow(_everythingService, _trayRepository, _explorerWindowService, _preferencesRepository);
        _dialogWatcher = new WinEventFileDialogWatcher(Dispatcher);
        _dialogWatcher.ActiveDialogChanged += OnActiveDialogChanged;
        _dialogWatcher.ActiveDialogMoved += OnActiveDialogMoved;
        _dialogWatcher.ActiveDialogContentChanged += OnActiveDialogContentChanged;
        _dialogWatcher.Start();
        _dialogDetachDebounceTimer = new DispatcherTimer(DispatcherPriority.Input)
        {
            Interval = DialogDetachDebounceInterval,
        };
        _dialogDetachDebounceTimer.Tick += DialogDetachDebounceTimer_OnTick;

        _trayIcon = new AppTrayIcon(LoadAppIcon());
        _trayIcon.OpenRequested += OnTrayOpenRequested;
        _trayIcon.ExitRequested += OnTrayExitRequested;

        SystemEvents.UserPreferenceChanged += OnSystemUserPreferenceChanged;
        ApplySystemTheme();

        if (_preferencesRepository.MainWindowHotkeyEnabled)
        {
            InitializeGlobalHotkey();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _isExitRequested = true;
        DispatcherUnhandledException -= OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException -= OnAppDomainUnhandledException;
        TaskScheduler.UnobservedTaskException -= OnTaskSchedulerUnobservedTaskException;
        SystemEvents.UserPreferenceChanged -= OnSystemUserPreferenceChanged;

        DisposeTrayIcon();
        DisposeGlobalHotkey();
        DisposeDialogWatcher();
        DisposeDialogDetachTimer();
        DisposeOverlayWindow();
        DisposeExplorerWindowService();
        DisposeMainWindow(closeWindow: false);

        AppLogger.LogInfo("Application exit.");
        base.OnExit(e);
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        if (ReferenceEquals(MainWindow, sender))
        {
            MainWindow = null;
        }

        RequestMemoryTrim("MainWindowClosed");
    }

    private void OnActiveDialogChanged(IntPtr? hwnd)
    {
        if (_isExitRequested)
        {
            return;
        }

        if (_overlayWindow is null)
        {
            return;
        }

        if (hwnd is { } target && target != IntPtr.Zero)
        {
            _dialogDetachDebounceTimer?.Stop();
            if (_activeDialogHwnd != target)
            {
                _activeDialogHwnd = target;
                _overlayWindow.AttachToDialog(target);
            }

            return;
        }

        if (_activeDialogHwnd == IntPtr.Zero)
        {
            return;
        }

        if (_dialogDetachDebounceTimer is not null && !_dialogDetachDebounceTimer.IsEnabled)
        {
            _dialogDetachDebounceTimer.Start();
        }
    }

    private void OnActiveDialogMoved(IntPtr hwnd)
    {
        if (_isExitRequested)
        {
            return;
        }

        if (_activeDialogHwnd == IntPtr.Zero || hwnd != _activeDialogHwnd)
        {
            return;
        }

        _overlayWindow?.RepositionToDialog();
    }

    private void OnActiveDialogContentChanged(IntPtr hwnd)
    {
        if (_isExitRequested || _overlayWindow is null)
        {
            return;
        }

        if (_activeDialogHwnd == IntPtr.Zero || hwnd != _activeDialogHwnd)
        {
            return;
        }

        _overlayWindow.OnDialogContentChanged();
    }

    private void DialogDetachDebounceTimer_OnTick(object? sender, EventArgs e)
    {
        _dialogDetachDebounceTimer?.Stop();
        if (_isExitRequested || _overlayWindow is null || _activeDialogHwnd == IntPtr.Zero)
        {
            return;
        }

        _activeDialogHwnd = IntPtr.Zero;
        _overlayWindow.DetachDialog();
        if (MainWindow is null)
        {
            RequestMemoryTrim("DialogDetached");
        }
    }

    private void OnTrayOpenRequested()
    {
        _ = Dispatcher.BeginInvoke(new Action(ShowMainWindow));
    }

    private void OnTrayExitRequested()
    {
        _ = Dispatcher.BeginInvoke(new Action(ExitApplication));
    }

    private void ShowMainWindow()
    {
        if (_isExitRequested)
        {
            return;
        }

        if (_dialogWatcher is null || _everythingService is null || _trayRepository is null || _preferencesRepository is null)
        {
            return;
        }

        var window = EnsureMainWindow(_dialogWatcher, _everythingService, _trayRepository, _preferencesRepository);

        if (!window.IsVisible)
        {
            window.Show();
        }

        if (window.WindowState == WindowState.Minimized)
        {
            window.WindowState = WindowState.Normal;
        }

        _ = window.Activate();
    }

    private MainWindow EnsureMainWindow(
        WinEventFileDialogWatcher watcher,
        EverythingService everythingService,
        TrayRepository trayRepository,
        UserPreferencesRepository preferencesRepository)
    {
        if (MainWindow is MainWindow existing)
        {
            return existing;
        }

        var created = new MainWindow(watcher, everythingService, trayRepository, preferencesRepository);
        created.ApplyTheme(_isSystemDarkMode);
        created.Closed += MainWindow_OnClosed;
        MainWindow = created;
        return created;
    }

    private void OnSystemUserPreferenceChanged(object? sender, UserPreferenceChangedEventArgs e)
    {
        if (_isExitRequested)
        {
            return;
        }

        if (e.Category is not UserPreferenceCategory.Color and not UserPreferenceCategory.General and not UserPreferenceCategory.VisualStyle)
        {
            return;
        }

        _ = Dispatcher.BeginInvoke(new Action(ApplySystemTheme));
    }

    private void ApplySystemTheme()
    {
        var darkMode = SystemThemeHelper.IsSystemAppDarkMode();
        _isSystemDarkMode = darkMode;
        _overlayWindow?.ApplyTheme(darkMode);
        if (MainWindow is MainWindow mainWindow)
        {
            mainWindow.ApplyTheme(darkMode);
        }
    }

    private void ExitApplication()
    {
        _isExitRequested = true;
        DisposeTrayIcon();
        DisposeGlobalHotkey();
        DisposeDialogWatcher();
        DisposeDialogDetachTimer();
        DisposeOverlayWindow();
        DisposeExplorerWindowService();
        DisposeMainWindow(closeWindow: true);

        Shutdown();
    }

    private void DisposeTrayIcon()
    {
        if (_trayIcon is null)
        {
            return;
        }

        _trayIcon.OpenRequested -= OnTrayOpenRequested;
        _trayIcon.ExitRequested -= OnTrayExitRequested;
        _trayIcon.Dispose();
        _trayIcon = null;
    }

    private void DisposeDialogWatcher()
    {
        if (_dialogWatcher is null)
        {
            return;
        }

        _dialogWatcher.ActiveDialogChanged -= OnActiveDialogChanged;
        _dialogWatcher.ActiveDialogMoved -= OnActiveDialogMoved;
        _dialogWatcher.ActiveDialogContentChanged -= OnActiveDialogContentChanged;
        _dialogWatcher.Dispose();
        _dialogWatcher = null;
    }

    private void DisposeDialogDetachTimer()
    {
        if (_dialogDetachDebounceTimer is null)
        {
            return;
        }

        _dialogDetachDebounceTimer.Tick -= DialogDetachDebounceTimer_OnTick;
        _dialogDetachDebounceTimer.Stop();
        _dialogDetachDebounceTimer = null;
    }

    private void DisposeOverlayWindow()
    {
        if (_overlayWindow is null)
        {
            return;
        }

        _overlayWindow.Shutdown();
        _overlayWindow = null;
    }

    private void DisposeExplorerWindowService()
    {
        _explorerWindowService?.Dispose();
        _explorerWindowService = null;
    }

    private void DisposeMainWindow(bool closeWindow)
    {
        if (MainWindow is not Window mainWindow)
        {
            return;
        }

        mainWindow.Closed -= MainWindow_OnClosed;
        if (closeWindow)
        {
            mainWindow.Close();
        }

        MainWindow = null;
    }

    private void InitializeGlobalHotkey()
    {
        if (_hotkeySource is not null)
        {
            return;
        }

        var parameters = new HwndSourceParameters("SuperSelect.HotkeySink")
        {
            Width = 0,
            Height = 0,
            WindowStyle = unchecked((int)0x80000000), // WS_POPUP
        };

        _hotkeySource = new HwndSource(parameters);
        _hotkeySource.AddHook(HotkeyWndProc);

        if (!Native.NativeMethods.RegisterHotKey(
                _hotkeySource.Handle,
                MainWindowHotkeyId,
                MainWindowHotkeyModifiers,
                MainWindowHotkeyVKey))
        {
            var error = Marshal.GetLastWin32Error();
            AppLogger.LogWarning($"RegisterHotKey failed. hotkey={MainWindowHotkeyDisplayName}; win32={error}");
            return;
        }

        AppLogger.LogInfo($"Global hotkey registered: {MainWindowHotkeyDisplayName}");
    }

    internal void SetMainWindowHotkeyEnabled(bool enabled)
    {
        if (_isExitRequested)
        {
            return;
        }

        if (enabled)
        {
            InitializeGlobalHotkey();
            return;
        }

        DisposeGlobalHotkey();
    }

    private void DisposeGlobalHotkey()
    {
        if (_hotkeySource is null)
        {
            return;
        }

        try
        {
            if (_hotkeySource.Handle != IntPtr.Zero &&
                !Native.NativeMethods.UnregisterHotKey(_hotkeySource.Handle, MainWindowHotkeyId))
            {
                var error = Marshal.GetLastWin32Error();
                if (error != 0)
                {
                    AppLogger.LogWarning($"UnregisterHotKey failed. hotkey={MainWindowHotkeyDisplayName}; win32={error}");
                }
            }

            _hotkeySource.RemoveHook(HotkeyWndProc);
            _hotkeySource.Dispose();
            _hotkeySource = null;
        }
        catch (Exception ex)
        {
            AppLogger.LogException("App.DisposeGlobalHotkey", ex, throttle: TimeSpan.FromSeconds(3));
        }
    }

    private IntPtr HotkeyWndProc(
        IntPtr hwnd,
        int msg,
        IntPtr wParam,
        IntPtr lParam,
        ref bool handled)
    {
        if (msg != Native.NativeMethods.WM_HOTKEY || wParam.ToInt32() != MainWindowHotkeyId)
        {
            return IntPtr.Zero;
        }

        _ = Dispatcher.BeginInvoke(new Action(ShowMainWindow));
        handled = true;
        return IntPtr.Zero;
    }

    private void RequestMemoryTrim(string reason)
    {
        if (_isExitRequested)
        {
            return;
        }

        var nowUtc = DateTime.UtcNow;
        if (nowUtc - _lastMemoryTrimUtc < MemoryTrimCooldown)
        {
            return;
        }

        if (GC.GetTotalMemory(forceFullCollection: false) < MemoryTrimManagedThresholdBytes)
        {
            return;
        }

        _lastMemoryTrimUtc = nowUtc;
        _ = Task.Run(() => TrimProcessMemory(reason));
    }

    private static void TrimProcessMemory(string reason)
    {
        try
        {
            // Let the detach/hide rendering complete before any global GC pause.
            Task.Delay(MemoryTrimStartDelay).GetAwaiter().GetResult();

            GC.Collect(GC.MaxGeneration, GCCollectionMode.Optimized, blocking: false, compacting: false);

            if (!string.Equals(reason, DialogDetachedTrimReason, StringComparison.Ordinal))
            {
                _ = Native.NativeMethods.EmptyWorkingSet(Native.NativeMethods.GetCurrentProcess());
            }

            AppLogger.LogInfo($"Background memory trim executed: {reason}");
        }
        catch (Exception ex)
        {
            AppLogger.LogException("App.TrimProcessMemory", ex, throttle: TimeSpan.FromSeconds(5));
        }
    }

    private static Icon LoadAppIcon()
    {
        try
        {
            var uri = new Uri("pack://application:,,,/Resources/app.ico", UriKind.Absolute);
            var streamInfo = System.Windows.Application.GetResourceStream(uri);
            if (streamInfo?.Stream is not null)
            {
                using var stream = streamInfo.Stream;
                using var memory = new MemoryStream();
                stream.CopyTo(memory);
                memory.Position = 0;
                using var icon = new Icon(memory);
                return (Icon)icon.Clone();
            }
        }
        catch
        {
            // Fall back to system icon.
        }

        return SystemIcons.Application;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppLogger.LogException(
            "DispatcherUnhandledException",
            e.Exception,
            throttle: TimeSpan.FromSeconds(3));

        if (!IsTransientAutomationException(e.Exception))
        {
            return;
        }

        e.Handled = true;
        _overlayWindow?.DetachDialog();
    }

    private void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            AppLogger.LogException("AppDomain.CurrentDomain.UnhandledException", ex);
        }
        else
        {
            AppLogger.LogWarning($"Unhandled non-Exception object: {e.ExceptionObject}");
        }
    }

    private void OnTaskSchedulerUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        AppLogger.LogException(
            "TaskScheduler.UnobservedTaskException",
            e.Exception,
            throttle: TimeSpan.FromSeconds(3));
        e.SetObserved();
    }

    private static bool IsTransientAutomationException(Exception exception)
    {
        Exception? current = exception;
        while (current is not null)
        {
            if (current is ElementNotAvailableException)
            {
                return true;
            }

            if (current is COMException com &&
                (com.HResult == unchecked((int)0x80010108) || // RPC_E_DISCONNECTED
                 com.HResult == unchecked((int)0x80010012)))   // RPC_E_SERVER_DIED_DNE
            {
                return true;
            }

            current = current.InnerException;
        }

        return false;
    }
}
