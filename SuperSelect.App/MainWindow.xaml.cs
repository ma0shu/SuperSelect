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
    private bool _isUpdatingAdminStartupUi;

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

        var isAdmin = AdminPrivilegeHelper.IsRunAsAdministrator();
        var isConfigured = AdminPrivilegeHelper.CheckRunOnStartupConfigured();

        AdminWarningText.Visibility = isAdmin ? Visibility.Collapsed : Visibility.Visible;

        _isUpdatingAdminStartupUi = true;
        AdminStartupCheckBox.IsChecked = isConfigured;
        AdminStartupCheckBox.Content = isConfigured
            ? "开机自启并提权（已启用）"
            : "开机自启并提权（免弹窗）";
        _isUpdatingAdminStartupUi = false;

        EverythingStatusText.Text = _everythingService.IsAvailable
            ? $"Everything 状态：可用（需要 Everything.exe 运行，并可加载 Everything64.dll）{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}"
            : $"Everything 状态：不可用 - {_everythingService.LastErrorMessage}{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}";

        RefreshTrayList();
    }

    private void AdminStartupCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        ChangeAdminStartupSetting(enable: true);
    }

    private void AdminStartupCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
    {
        ChangeAdminStartupSetting(enable: false);
    }

    private void ChangeAdminStartupSetting(bool enable)
    {
        if (_isUpdatingAdminStartupUi)
        {
            return;
        }

        try
        {
            if (!enable)
            {
                AdminPrivilegeHelper.DisableRunOnStartupAsAdmin();
            }
            else
            {
                AdminPrivilegeHelper.EnableRunOnStartupAsAdmin();

                // Startup task takes effect on next launch. Restart now if current session is not elevated.
                PromptRestartIfElevationNeeded();
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, $"设置失败，可能被拦截或取消了授权：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        RefreshUi();
    }

    private void PromptRestartIfElevationNeeded()
    {
        if (AdminPrivilegeHelper.IsRunAsAdministrator())
        {
            return;
        }

        var res = System.Windows.MessageBox.Show(
            this,
            "已完成开机自启提权配置。\n当前进程仍是普通权限，重启后才会切到管理员权限运行。\n\n是否立即重启并提权？",
            "建议重启",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (res != MessageBoxResult.Yes)
        {
            return;
        }

        AdminPrivilegeHelper.RestartCurrentProcessAsAdministrator();
        System.Windows.Application.Current.Shutdown();
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
