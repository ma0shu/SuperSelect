using System.ComponentModel;
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

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
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

        var mainWindow = new MainWindow(_dialogWatcher, _everythingService, _trayRepository);
        mainWindow.Closing += MainWindow_OnClosing;
        MainWindow = mainWindow;
        mainWindow.Show();

        _trayIcon = new AppTrayIcon(LoadAppIcon());
        _trayIcon.OpenRequested += OnTrayOpenRequested;
        _trayIcon.ExitRequested += OnTrayExitRequested;
    }

    protected override void OnExit(ExitEventArgs e)
    {
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
            _overlayWindow.Close();
        }

        AppLogger.LogInfo("Application exit.");
        base.OnExit(e);
    }

    private void MainWindow_OnClosing(object? sender, CancelEventArgs e)
    {
        if (_isExitRequested || MainWindow is null)
        {
            return;
        }

        e.Cancel = true;
        MainWindow.Hide();
    }

    private void OnActiveDialogChanged(IntPtr? hwnd)
    {
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
    }

    private void OnActiveDialogMoved(IntPtr hwnd)
    {
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
        if (MainWindow is not Window window)
        {
            return;
        }

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

    private void ExitApplication()
    {
        _isExitRequested = true;
        if (MainWindow is not null)
        {
            MainWindow.Closing -= MainWindow_OnClosing;
            MainWindow.Close();
        }

        Shutdown();
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
