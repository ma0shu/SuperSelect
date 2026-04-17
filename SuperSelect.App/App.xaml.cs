using System.Windows;
using SuperSelect.App.Services;
using SuperSelect.App.Views;

namespace SuperSelect.App;

public partial class App : Application
{
    private WinEventFileDialogWatcher? _dialogWatcher;
    private OverlayWindow? _overlayWindow;

    private EverythingService? _everythingService;
    private TrayRepository? _trayRepository;
    private ExplorerWindowService? _explorerWindowService;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _everythingService = new EverythingService();
        _trayRepository = new TrayRepository();
        _explorerWindowService = new ExplorerWindowService();

        _overlayWindow = new OverlayWindow(_everythingService, _trayRepository, _explorerWindowService);
        _dialogWatcher = new WinEventFileDialogWatcher(Dispatcher);
        _dialogWatcher.ActiveDialogChanged += OnActiveDialogChanged;
        _dialogWatcher.Start();

        var mainWindow = new MainWindow(_dialogWatcher, _everythingService, _trayRepository);
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        if (_dialogWatcher is not null)
        {
            _dialogWatcher.ActiveDialogChanged -= OnActiveDialogChanged;
            _dialogWatcher.Dispose();
        }

        if (_overlayWindow is not null)
        {
            _overlayWindow.Close();
        }

        base.OnExit(e);
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
}
