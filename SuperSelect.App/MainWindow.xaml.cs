using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Media;
using SuperSelect.App.Models;
using SuperSelect.App.Services;
using MediaColor = System.Windows.Media.Color;

namespace SuperSelect.App;

public partial class MainWindow : Window
{
    private sealed class BlockedRecentItem
    {
        public required string FullPath { get; init; }
        public required string DisplayName { get; init; }
        public string SecondaryText { get; init; } = string.Empty;
        public bool IsDirectory { get; init; }
        public string TypeLabel => IsDirectory ? "目录" : "文件";
    }

    private readonly WinEventFileDialogWatcher _watcher;
    private readonly EverythingService _everythingService;
    private readonly TrayRepository _trayRepository;
    private readonly UserPreferencesRepository _preferencesRepository;
    private readonly ObservableCollection<FileCandidate> _trayItems = [];
    private readonly ObservableCollection<BlockedRecentItem> _blockedRecentItems = [];
    private bool _isUpdatingUi;

    internal MainWindow(
        WinEventFileDialogWatcher watcher,
        EverythingService everythingService,
        TrayRepository trayRepository,
        UserPreferencesRepository preferencesRepository)
    {
        InitializeComponent();

        _watcher = watcher;
        _everythingService = everythingService;
        _trayRepository = trayRepository;
        _preferencesRepository = preferencesRepository;

        TrayList.ItemsSource = _trayItems;
        BlockedRecentList.ItemsSource = _blockedRecentItems;
        RefreshUi();
    }

    internal void ApplyTheme(bool isDarkMode)
    {
        if (isDarkMode)
        {
            SetThemeBrush("MainWindowWindowBackgroundBrush", MediaColor.FromRgb(0x1E, 0x1E, 0x1E));
            SetThemeBrush("MainWindowPanelBackgroundBrush", MediaColor.FromRgb(0x25, 0x25, 0x26));
            SetThemeBrush("MainWindowListBackgroundBrush", MediaColor.FromRgb(0x25, 0x25, 0x26));
            SetThemeBrush("MainWindowBorderBrush", MediaColor.FromRgb(0x3F, 0x3F, 0x46));
            SetThemeBrush("MainWindowPrimaryTextBrush", MediaColor.FromRgb(0xE5, 0xE7, 0xEB));
            SetThemeBrush("MainWindowSecondaryTextBrush", MediaColor.FromRgb(0xA1, 0xA1, 0xAA));
            SetThemeBrush("MainWindowStatusTextBrush", MediaColor.FromRgb(0xC4, 0xC4, 0xCC));
            SetThemeBrush("MainWindowWarningBrush", MediaColor.FromRgb(0xF8, 0x71, 0x71));
            SetThemeBrush("MainWindowAccentBrush", MediaColor.FromRgb(0x60, 0xA5, 0xFA));
            SetThemeBrush("MainWindowDangerBrush", MediaColor.FromRgb(0xF8, 0x71, 0x71));
            SetThemeBrush("MainWindowItemIconBrush", MediaColor.FromRgb(0xA1, 0xA1, 0xAA));
            SetThemeBrush("MainWindowItemNameBrush", MediaColor.FromRgb(0xE5, 0xE7, 0xEB));
            SetThemeBrush("MainWindowItemSecondaryBrush", MediaColor.FromRgb(0x9C, 0xA3, 0xAF));
            SetThemeBrush("MainWindowItemHoverBrush", MediaColor.FromRgb(0x2F, 0x31, 0x36));
            SetThemeBrush("MainWindowItemSelectedBrush", MediaColor.FromRgb(0x1E, 0x3A, 0x5F));
            SetThemeBrush("MainWindowFolderIconBrush", MediaColor.FromRgb(0xFB, 0xBF, 0x24));
            return;
        }

        SetThemeBrush("MainWindowWindowBackgroundBrush", MediaColor.FromRgb(0xF4, 0xF6, 0xF8));
        SetThemeBrush("MainWindowPanelBackgroundBrush", MediaColor.FromRgb(0xFF, 0xFF, 0xFF));
        SetThemeBrush("MainWindowListBackgroundBrush", MediaColor.FromRgb(0xFF, 0xFF, 0xFF));
        SetThemeBrush("MainWindowBorderBrush", MediaColor.FromRgb(0xE2, 0xE8, 0xF0));
        SetThemeBrush("MainWindowPrimaryTextBrush", MediaColor.FromRgb(0x33, 0x41, 0x55));
        SetThemeBrush("MainWindowSecondaryTextBrush", MediaColor.FromRgb(0x64, 0x74, 0x8B));
        SetThemeBrush("MainWindowStatusTextBrush", MediaColor.FromRgb(0x47, 0x55, 0x69));
        SetThemeBrush("MainWindowWarningBrush", MediaColor.FromRgb(0xEF, 0x44, 0x44));
        SetThemeBrush("MainWindowAccentBrush", MediaColor.FromRgb(0x25, 0x63, 0xEB));
        SetThemeBrush("MainWindowDangerBrush", MediaColor.FromRgb(0xEF, 0x44, 0x44));
        SetThemeBrush("MainWindowItemIconBrush", MediaColor.FromRgb(0x94, 0xA3, 0xB8));
        SetThemeBrush("MainWindowItemNameBrush", MediaColor.FromRgb(0x1E, 0x29, 0x3B));
        SetThemeBrush("MainWindowItemSecondaryBrush", MediaColor.FromRgb(0x94, 0xA3, 0xB8));
        SetThemeBrush("MainWindowItemHoverBrush", MediaColor.FromRgb(0xF8, 0xFA, 0xFC));
        SetThemeBrush("MainWindowItemSelectedBrush", MediaColor.FromRgb(0xEF, 0xF6, 0xFF));
        SetThemeBrush("MainWindowFolderIconBrush", MediaColor.FromRgb(0xFC, 0xD3, 0x4D));
    }

    private void SetThemeBrush(string key, MediaColor color)
    {
        if (Resources[key] is SolidColorBrush brush)
        {
            if (!brush.IsFrozen)
            {
                brush.Color = color;
                return;
            }
        }

        Resources[key] = new SolidColorBrush(color);
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

        _isUpdatingUi = true;
        AdminStartupCheckBox.IsChecked = isConfigured;
        AdminStartupCheckBox.Content = isConfigured
            ? "开机自启并提权（已启用）"
            : "开机自启并提权（免弹窗）";
        MainWindowHotkeyCheckBox.IsChecked = _preferencesRepository.MainWindowHotkeyEnabled;
        AutoJumpExplorerCheckBox.IsChecked = _preferencesRepository.AutoJumpToExplorerOnAttachEnabled;
        _isUpdatingUi = false;

        EverythingStatusText.Text = _everythingService.IsAvailable
            ? $"Everything 状态：可用（需要 Everything.exe 运行，并可加载 Everything64.dll）{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}"
            : $"Everything 状态：不可用 - {_everythingService.LastErrorMessage}{Environment.NewLine}日志目录：{AppLogger.LogDirectoryPath}";

        RefreshTrayList();
        RefreshBlockedRecentList();
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
        if (_isUpdatingUi)
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

    private void AutoJumpExplorerCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        ChangeAutoJumpExplorerSetting(enable: true);
    }

    private void AutoJumpExplorerCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
    {
        ChangeAutoJumpExplorerSetting(enable: false);
    }

    private void ChangeAutoJumpExplorerSetting(bool enable)
    {
        if (_isUpdatingUi)
        {
            return;
        }

        _preferencesRepository.SetAutoJumpToExplorerOnAttachEnabled(enable);
    }

    private void MainWindowHotkeyCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        ChangeMainWindowHotkeySetting(enable: true);
    }

    private void MainWindowHotkeyCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
    {
        ChangeMainWindowHotkeySetting(enable: false);
    }

    private void ChangeMainWindowHotkeySetting(bool enable)
    {
        if (_isUpdatingUi)
        {
            return;
        }

        _preferencesRepository.SetMainWindowHotkeyEnabled(enable);
        if (System.Windows.Application.Current is App app)
        {
            app.SetMainWindowHotkeyEnabled(enable);
        }
    }

    private void PromptRestartIfElevationNeeded()
    {
        if (AdminPrivilegeHelper.IsRunAsAdministrator())
        {
            return;
        }

        var res = System.Windows.MessageBox.Show(
            this,
            "已完成开机自启提权配置。\n当前进程仍是普通权限，重启软件后才会切到管理员权限运行。\n\n是否立即重启进程并提权？",
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
            _trayItems.Add(candidate);
        }
    }

    private void RefreshBlockedRecentList()
    {
        _blockedRecentItems.Clear();
        foreach (var entry in _everythingService.GetRecentBlockEntries())
        {
            var pathForDisplay = entry.IsDirectory
                ? Path.TrimEndingDirectorySeparator(entry.FullPath)
                : entry.FullPath;
            var displayName = Path.GetFileName(pathForDisplay);
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = pathForDisplay;
            }

            _blockedRecentItems.Add(new BlockedRecentItem
            {
                FullPath = entry.FullPath,
                DisplayName = displayName,
                SecondaryText = Path.GetDirectoryName(pathForDisplay) ?? string.Empty,
                IsDirectory = entry.IsDirectory,
            });
        }
    }

    private void RefreshButton_OnClick(object sender, RoutedEventArgs e)
    {
        RefreshUi();
    }

    private void RefreshBlockedButton_OnClick(object sender, RoutedEventArgs e)
    {
        RefreshBlockedRecentList();
    }

    private void RemoveBlockedButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (BlockedRecentList.SelectedItem is not BlockedRecentItem item)
        {
            return;
        }

        _ = item.IsDirectory
            ? _everythingService.UnblockRecentDirectory(item.FullPath)
            : _everythingService.UnblockRecentFile(item.FullPath);
        RefreshBlockedRecentList();
    }

    private void RemoveButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (TrayList.SelectedItem is not FileCandidate candidate)
        {
            return;
        }

        _trayRepository.Remove(candidate.FullPath);
        RefreshTrayList();
    }

    private void ClearButton_OnClick(object sender, RoutedEventArgs e)
    {
        _trayRepository.Clear();
        RefreshTrayList();
    }

    private void TrayDropHost_OnPreviewDragOver(object sender, System.Windows.DragEventArgs e)
    {
        e.Effects = DropPathExtractor.HasDroppedFiles(e.Data)
            ? System.Windows.DragDropEffects.Copy
            : System.Windows.DragDropEffects.None;
        e.Handled = true;
    }

    private void TrayDropHost_OnPreviewDrop(object sender, System.Windows.DragEventArgs e)
    {
        var paths = DropPathExtractor.ExtractDroppedPaths(e.Data);
        if (paths.Count == 0)
        {
            e.Handled = true;
            return;
        }

        _ = _trayRepository.AddMany(paths);
        RefreshTrayList();
        e.Handled = true;
    }
}
