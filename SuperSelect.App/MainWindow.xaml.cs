using System.Collections.ObjectModel;
using System.Windows;
using SuperSelect.App.Services;

namespace SuperSelect.App;

public partial class MainWindow : Window
{
    private readonly WinEventFileDialogWatcher _watcher;
    private readonly EverythingService _everythingService;
    private readonly TrayRepository _trayRepository;
    private readonly ObservableCollection<string> _trayItems = [];

    internal MainWindow(
        WinEventFileDialogWatcher watcher,
        EverythingService everythingService,
        TrayRepository trayRepository)
    {
        InitializeComponent();

        _watcher = watcher;
        _everythingService = everythingService;
        _trayRepository = trayRepository;

        TrayList.ItemsSource = _trayItems;
        RefreshUi();
    }

    protected override void OnActivated(EventArgs e)
    {
        base.OnActivated(e);
        RefreshUi();
    }

    private void RefreshUi()
    {
        HookStatusText.Text = _watcher.IsRunning
            ? "Hook 状态：运行中"
            : "Hook 状态：未运行";

        EverythingStatusText.Text = _everythingService.IsAvailable
            ? $"Everything 状态：可用（需要 Everything.exe 运行，并可加载 Everything64.dll）{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}"
            : $"Everything 状态：不可用 - {_everythingService.LastErrorMessage}{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}";

        RefreshTrayList();
    }

    private void RefreshTrayList()
    {
        _trayItems.Clear();
        foreach (var candidate in _trayRepository.Query(string.Empty))
        {
            _trayItems.Add(candidate.FullPath);
        }
    }

    private void RefreshButton_OnClick(object sender, RoutedEventArgs e)
    {
        RefreshUi();
    }

    private void RemoveButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (TrayList.SelectedItem is not string path)
        {
            return;
        }

        _trayRepository.Remove(path);
        RefreshTrayList();
    }

    private void ClearButton_OnClick(object sender, RoutedEventArgs e)
    {
        _trayRepository.Clear();
        RefreshTrayList();
    }
}
