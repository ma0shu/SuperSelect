using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Threading;
using SuperSelect.App.Services;
using SuperSelect.App.Views;

namespace SuperSelect.App;

public partial class App : System.Windows.Application
{
    private WinEventFileDialogWatcher? _dialogWatcher;
    private OverlayWindow? _overlayWindow;
    private AppTrayIcon? _trayIcon;
    private bool _isExitRequested;

    private EverythingService? _everythingService;
    private TrayRepository? _trayRepository;
    private ExplorerWindowService? _explorerWindowService;
    private UserPreferencesRepository? _preferencesRepository;
    private DateTime _lastMemoryTrimUtc = DateTime.MinValue;
    private static readonly TimeSpan MemoryTrimCooldown = TimeSpan.FromMinutes(2);
    private const long MemoryTrimManagedThresholdBytes = 64L * 1024 * 1024;

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
        _dialogWatcher.Start();

        _trayIcon = new AppTrayIcon(LoadAppIcon());
        _trayIcon.OpenRequested += OnTrayOpenRequested;
        _trayIcon.ExitRequested += OnTrayExitRequested;
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _isExitRequested = true;
        DispatcherUnhandledException -= OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException -= OnAppDomainUnhandledException;
        TaskScheduler.UnobservedTaskException -= OnTaskSchedulerUnobservedTaskException;

        if (_trayIcon is not null)
        {
            _trayIcon.OpenRequested -= OnTrayOpenRequested;
            _trayIcon.ExitRequested -= OnTrayExitRequested;
            _trayIcon.Dispose();
            _trayIcon = null;
        }

        if (_dialogWatcher is not null)
        {
            _dialogWatcher.ActiveDialogChanged -= OnActiveDialogChanged;
            _dialogWatcher.ActiveDialogMoved -= OnActiveDialogMoved;
            _dialogWatcher.Dispose();
        }

        if (_overlayWindow is not null)
        {
            _overlayWindow.Shutdown();
            _overlayWindow = null;
        }

        if (MainWindow is Window mainWindow)
        {
            mainWindow.Closed -= MainWindow_OnClosed;
            MainWindow = null;
        }

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
            _overlayWindow.AttachToDialog(target);
            return;
        }

        _overlayWindow.DetachDialog();
        if (MainWindow is null)
        {
            RequestMemoryTrim("DialogDetached");
        }
    }

    private void OnActiveDialogMoved(IntPtr hwnd)
    {
        if (_isExitRequested)
        {
            return;
        }

        _overlayWindow?.RepositionToDialog();
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

        if (_dialogWatcher is null || _everythingService is null || _trayRepository is null)
        {
            return;
        }

        var window = EnsureMainWindow(_dialogWatcher, _everythingService, _trayRepository);

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
        TrayRepository trayRepository)
    {
        if (MainWindow is MainWindow existing)
        {
            return existing;
        }

        var created = new MainWindow(watcher, everythingService, trayRepository);
        created.Closed += MainWindow_OnClosed;
        MainWindow = created;
        return created;
    }

    private void ExitApplication()
    {
        _isExitRequested = true;

        if (_dialogWatcher is not null)
        {
            _dialogWatcher.ActiveDialogChanged -= OnActiveDialogChanged;
            _dialogWatcher.ActiveDialogMoved -= OnActiveDialogMoved;
            _dialogWatcher.Dispose();
            _dialogWatcher = null;
        }

        if (_overlayWindow is not null)
        {
            _overlayWindow.Shutdown();
            _overlayWindow = null;
        }

        if (MainWindow is not null)
        {
            MainWindow.Closed -= MainWindow_OnClosed;
            MainWindow.Close();
            MainWindow = null;
        }

        Shutdown();
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
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Optimized, blocking: false, compacting: false);

            _ = Native.NativeMethods.EmptyWorkingSet(Native.NativeMethods.GetCurrentProcess());
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
