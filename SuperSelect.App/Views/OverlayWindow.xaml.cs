using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using SuperSelect.App.Models;
using SuperSelect.App.Native;
using SuperSelect.App.Services;

namespace SuperSelect.App.Views;

public partial class OverlayWindow : Window
{
    private sealed record OptionItem<T>(string Label, T Value);

    private const double CompactHeightDip = 108.0;
    private const double ExpandedHeightDip = 350.0;
    private const double MinUsableHeightDip = 72.0;
    private const double HeaderAndStatusHeightDip = 112.0;
    private const int OverlayGapPx = 8;
    private const int QueryResultLimit = 5000;

    private readonly EverythingService _everythingService;
    private readonly TrayRepository _trayRepository;
    private readonly ExplorerWindowService _explorerWindowService;
    private readonly UserPreferencesRepository _preferencesRepository;
    private readonly ObservableCollection<FileCandidate> _items = [];
    private readonly DispatcherTimer _searchDebounceTimer;
    private readonly DispatcherTimer _dialogMonitorTimer;
    private readonly List<OptionItem<EverythingSortOption>> _sortOptions;

    private CancellationTokenSource? _refreshCts;
    private FileDialogAutomationController? _dialogController;
    private ScrollViewer? _resultListScrollViewer;
    private IntPtr _dialogHwnd;
    private int _wheelDeltaAccumulator;
    private bool _suppressSingleSelect;
    private OverlayMode _currentMode = OverlayMode.Search;
    private int _sortIndex;
    private bool _preferredTypeFilterEnabled;
    private bool _isFolderDialog;
    private string _lastFileTypeFilterText = string.Empty;

    internal OverlayWindow(
        EverythingService everythingService,
        TrayRepository trayRepository,
        ExplorerWindowService explorerWindowService,
        UserPreferencesRepository preferencesRepository)
    {
        InitializeComponent();

        _everythingService = everythingService;
        _trayRepository = trayRepository;
        _explorerWindowService = explorerWindowService;
        _preferencesRepository = preferencesRepository;
        _preferredTypeFilterEnabled = _preferencesRepository.TypeFilterEnabled;

        _sortOptions =
        [
            new OptionItem<EverythingSortOption>("时间↓", EverythingSortOption.DateModifiedDesc),
            new OptionItem<EverythingSortOption>("时间↑", EverythingSortOption.DateModifiedAsc),
            new OptionItem<EverythingSortOption>("名称A-Z", EverythingSortOption.NameAsc),
            new OptionItem<EverythingSortOption>("名称Z-A", EverythingSortOption.NameDesc),
            new OptionItem<EverythingSortOption>("路径A-Z", EverythingSortOption.PathAsc),
            new OptionItem<EverythingSortOption>("路径Z-A", EverythingSortOption.PathDesc),
        ];

        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(180),
        };
        _searchDebounceTimer.Tick += SearchDebounceTimer_OnTick;
        _dialogMonitorTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(420),
        };
        _dialogMonitorTimer.Tick += DialogMonitorTimer_OnTick;

        ResultList.ItemsSource = _items;
        ApplyModeUiState();
        UpdateSortButtonLabel();

        ShowActivated = false;
    }

    public void AttachToDialog(IntPtr dialogHwnd)
    {
        _dialogHwnd = dialogHwnd;
        _dialogController = new FileDialogAutomationController(dialogHwnd);
        _isFolderDialog = _dialogController.IsFolderSelectionDialog();
        _lastFileTypeFilterText = _dialogController.GetCurrentFileTypeFilterText();
        ApplyModeUiState();

        if (!IsVisible)
        {
            Show();
        }

        _dialogMonitorTimer.Start();
        PositionToDialog();
        _ = RefreshCandidatesAsync(immediate: true);
    }

    public void DetachDialog()
    {
        _dialogHwnd = IntPtr.Zero;
        _dialogController = null;
        _wheelDeltaAccumulator = 0;
        _dialogMonitorTimer.Stop();
        _isFolderDialog = false;
        _lastFileTypeFilterText = string.Empty;
        ApplyModeUiState();

        _refreshCts?.Cancel();
        _refreshCts = null;

        _items.Clear();
        ResultHost.Visibility = Visibility.Collapsed;
        Height = CompactHeightDip;

        if (IsVisible)
        {
            Hide();
        }
    }

    public void RepositionToDialog()
    {
        if (!IsVisible || _dialogHwnd == IntPtr.Zero)
        {
            return;
        }

        PositionToDialog();
    }

    private async Task RefreshCandidatesAsync(bool immediate = false)
    {
        if (_dialogController is null)
        {
            return;
        }

        _refreshCts?.Cancel();
        _refreshCts = new CancellationTokenSource();
        var token = _refreshCts.Token;

        var mode = _currentMode;
        var keyword = SearchBox.Text.Trim();
        var sort = CurrentSort;

        if (mode == OverlayMode.Search && string.IsNullOrWhiteSpace(keyword))
        {
            UpdateItems([]);
            SetStatus(_isFolderDialog ? "输入关键字开始搜索文件夹。" : "输入关键字开始搜索。");
            return;
        }

        if (!immediate)
        {
            await Task.Delay(1, token);
        }

        IReadOnlyList<FileCandidate> items = [];
        var statusMessage = string.Empty;
        var canApplyDialogTypeFilter = true;

        try
        {
            switch (mode)
            {
                case OverlayMode.Search:
                    items = _isFolderDialog
                        ? await _everythingService.SearchFoldersAsync(keyword, sort, QueryResultLimit, token)
                        : await _everythingService.SearchFilesAsync(keyword, sort, QueryResultLimit, token);
                    if (!_everythingService.IsAvailable)
                    {
                        statusMessage = $"Everything 不可用：{_everythingService.LastErrorMessage}";
                        canApplyDialogTypeFilter = false;
                    }
                    else
                    {
                        statusMessage = items.Count == 0
                            ? (_isFolderDialog ? "没有文件夹搜索结果。" : "没有搜索结果。")
                            : $"{(_isFolderDialog ? "文件夹搜索" : "搜索")}结果：{items.Count} 条。";
                    }

                    break;
                case OverlayMode.Recent:
                    items = _isFolderDialog
                        ? await _everythingService.RecentFoldersAsync(sort, QueryResultLimit, token)
                        : await _everythingService.RecentFilesAsync(sort, QueryResultLimit, token);
                    if (!_everythingService.IsAvailable)
                    {
                        statusMessage = $"Everything 不可用：{_everythingService.LastErrorMessage}";
                        canApplyDialogTypeFilter = false;
                    }
                    else
                    {
                        statusMessage = items.Count == 0
                            ? (_isFolderDialog ? "没有最近文件夹结果。" : "没有最近文件结果。")
                            : $"{(_isFolderDialog ? "最近文件夹" : "最近文件")}：{items.Count} 条。";
                    }

                    break;
                case OverlayMode.Tray:
                    items = _trayRepository.Query(string.Empty);
                    if (_isFolderDialog)
                    {
                        items = items.Where(candidate => candidate.IsDirectory).ToList();
                    }

                    statusMessage = items.Count == 0
                        ? (_isFolderDialog ? "托盘中没有可用文件夹，可直接拖入文件夹。"
                                           : "托盘为空，可直接把文件拖进面板。")
                        : $"{(_isFolderDialog ? "托盘文件夹" : "托盘文件")}：{items.Count} 条。";
                    break;
                case OverlayMode.Explorer:
                    items = _explorerWindowService.GetOpenLocations(string.Empty);
                    statusMessage = items.Count == 0
                        ? "未检测到可用资源管理器路径。"
                        : $"资源管理器路径：{items.Count} 条。";
                    canApplyDialogTypeFilter = false;
                    break;
            }
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (token.IsCancellationRequested)
        {
            return;
        }

        if (_isFolderDialog)
        {
            canApplyDialogTypeFilter = false;
        }

        if (IsTypeFilterActive && canApplyDialogTypeFilter)
        {
            var allowedExtensions = _dialogController.GetAllowedFileExtensions();
            if (allowedExtensions.Count > 0)
            {
                items = items
                    .Where(candidate => candidate.IsDirectory || HasAllowedExtension(candidate.FullPath, allowedExtensions))
                    .ToList();

                var caption = _currentMode switch
                {
                    OverlayMode.Search => "搜索",
                    OverlayMode.Recent => "最近",
                    OverlayMode.Tray => "托盘",
                    _ => "结果",
                };
                statusMessage = $"{caption}结果：{items.Count} 条（类型过滤）";
            }
            else if (!string.IsNullOrWhiteSpace(statusMessage))
            {
                statusMessage = $"{statusMessage}（当前类型=全部）";
            }
        }
        else if (_isFolderDialog && !string.IsNullOrWhiteSpace(statusMessage))
        {
            statusMessage = $"{statusMessage}（文件夹选择模式）";
        }

        if (!string.IsNullOrWhiteSpace(statusMessage))
        {
            SetStatus(statusMessage);
        }

        UpdateItems(items);
    }

    private void UpdateItems(IReadOnlyList<FileCandidate> items)
    {
        var previousVisibility = ResultHost.Visibility;
        var previousHeight = Height;
        var changed = !AreItemsEquivalent(items);

        if (changed)
        {
            _suppressSingleSelect = true;
            _items.Clear();
            foreach (var item in items)
            {
                _items.Add(item);
            }

            ResultList.SelectedItem = null;
            _suppressSingleSelect = false;
        }

        ResultHost.Visibility = _items.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        Height = _items.Count > 0 ? ExpandedHeightDip : CompactHeightDip;

        if (changed ||
            previousVisibility != ResultHost.Visibility ||
            Math.Abs(previousHeight - Height) > 0.5)
        {
            PositionToDialog();
        }
    }

    private void PositionToDialog()
    {
        if (_dialogHwnd == IntPtr.Zero || !NativeMethods.GetWindowRect(_dialogHwnd, out var rect))
        {
            return;
        }

        var dpi = NativeMethods.GetDpiForWindow(_dialogHwnd);
        if (dpi == 0)
        {
            dpi = 96;
        }

        var scale = dpi / 96.0;
        var minWidthPx = DipToPx(MinWidth, scale);
        var maxWidthPx = DipToPx(1400, scale);
        var desiredWidthPx = Math.Clamp(rect.Width, minWidthPx, maxWidthPx);

        var desiredHeightPx = DipToPx(Height, scale);
        var minUsableHeightPx = DipToPx(MinUsableHeightDip, scale);

        var monitor = NativeMethods.MonitorFromWindow(_dialogHwnd, NativeMethods.MONITOR_DEFAULTTONEAREST);
        var workArea = GetWorkArea(monitor, rect);

        var desiredLeftPx = rect.Left;
        if (desiredLeftPx + desiredWidthPx > workArea.Right)
        {
            desiredLeftPx = workArea.Right - desiredWidthPx;
        }

        desiredLeftPx = Math.Max(workArea.Left, desiredLeftPx);

        var desiredTopPx = rect.Bottom + OverlayGapPx;
        var maxTopPx = workArea.Bottom - minUsableHeightPx;
        if (desiredTopPx > maxTopPx)
        {
            desiredTopPx = maxTopPx;
        }

        desiredTopPx = Math.Max(workArea.Top, desiredTopPx);

        var availableHeightPx = workArea.Bottom - desiredTopPx;
        var finalHeightPx = Math.Max(1, Math.Min(desiredHeightPx, availableHeightPx));
        var finalHeightDip = PxToDip(finalHeightPx, scale);

        UpdateResultHostHeight(finalHeightDip);

        var overlayHwnd = new WindowInteropHelper(this).Handle;
        if (overlayHwnd != IntPtr.Zero)
        {
            _ = NativeMethods.SetWindowPos(
                overlayHwnd,
                IntPtr.Zero,
                desiredLeftPx,
                desiredTopPx,
                desiredWidthPx,
                finalHeightPx,
                NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_NOZORDER);
        }
        else
        {
            Width = PxToDip(desiredWidthPx, scale);
            Height = finalHeightDip;
            Left = PxToDip(desiredLeftPx, scale);
            Top = PxToDip(desiredTopPx, scale);
        }
    }

    private void UpdateResultHostHeight(double windowHeightDip)
    {
        if (ResultHost.Visibility != Visibility.Visible)
        {
            return;
        }

        var computed = windowHeightDip - HeaderAndStatusHeightDip;
        ResultHost.Height = Math.Max(ResultHost.MinHeight, computed);
    }

    private void SetStatus(string message)
    {
        StatusText.Text = message;
    }

    private EverythingSortOption CurrentSort =>
        _sortOptions[Math.Clamp(_sortIndex, 0, _sortOptions.Count - 1)].Value;

    private bool IsTypeFilterActive => _preferredTypeFilterEnabled && !_isFolderDialog;

    private async void SearchDebounceTimer_OnTick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        await RefreshCandidatesAsync();
    }

    private async void DialogMonitorTimer_OnTick(object? sender, EventArgs e)
    {
        try
        {
            if (_dialogController is null || _dialogHwnd == IntPtr.Zero)
            {
                return;
            }

            if (!NativeMethods.IsWindow(_dialogHwnd))
            {
                DetachDialog();
                return;
            }

            var folderMode = _dialogController.IsFolderSelectionDialog();
            var folderModeChanged = folderMode != _isFolderDialog;
            if (folderModeChanged)
            {
                _isFolderDialog = folderMode;
                ApplyModeUiState();
            }

            var filterChanged = false;
            if (IsTypeFilterActive && _currentMode != OverlayMode.Explorer)
            {
                var filterText = _dialogController.GetCurrentFileTypeFilterText();
                filterChanged = !string.Equals(filterText, _lastFileTypeFilterText, StringComparison.OrdinalIgnoreCase);
                if (filterChanged)
                {
                    _lastFileTypeFilterText = filterText;
                }
            }

            if (folderModeChanged || filterChanged)
            {
                _searchDebounceTimer.Stop();
                await RefreshCandidatesAsync(immediate: true);
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException(
                "OverlayWindow.DialogMonitorTimer_OnTick",
                ex,
                throttle: TimeSpan.FromSeconds(3));
            DetachDialog();
        }
    }

    private void SearchBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_currentMode != OverlayMode.Search)
        {
            return;
        }

        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void SearchBox_OnKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        var candidate = ResultList.SelectedItem as FileCandidate ?? _items.FirstOrDefault();
        if (candidate is null)
        {
            return;
        }

        ExecuteConfirm(candidate);
        e.Handled = true;
    }

    private void ResultList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSingleSelect || ResultList.SelectedItem is not FileCandidate candidate)
        {
            return;
        }

        ExecuteSingle(candidate);
    }

    private void ResultList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ResultList.SelectedItem is FileCandidate candidate)
        {
            ExecuteConfirm(candidate);
        }
    }

    private void ResultList_OnPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        var viewer = _resultListScrollViewer ??= FindDescendant<ScrollViewer>(ResultList);
        if (viewer is null)
        {
            return;
        }

        _wheelDeltaAccumulator += e.Delta;
        while (_wheelDeltaAccumulator >= 120)
        {
            viewer.LineUp();
            _wheelDeltaAccumulator -= 120;
        }

        while (_wheelDeltaAccumulator <= -120)
        {
            viewer.LineDown();
            _wheelDeltaAccumulator += 120;
        }

        e.Handled = true;
    }

    private void ResultList_OnKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != Key.Enter || ResultList.SelectedItem is not FileCandidate candidate)
        {
            return;
        }

        ExecuteConfirm(candidate);
        e.Handled = true;
    }

    private void ExecuteSingle(FileCandidate candidate)
    {
        if (_dialogController is null)
        {
            return;
        }

        var ok = _dialogController.TryPrimeSelection(candidate);
        if (!ok)
        {
            SetStatus("单击定位失败，目标对话框控件不可用。");
            return;
        }

        SetStatus($"已定位：{candidate.DisplayName}");
    }

    private void ExecuteConfirm(FileCandidate candidate)
    {
        if (_dialogController is null)
        {
            return;
        }

        var ok = _dialogController.TryConfirmSelection(candidate);
        if (!ok)
        {
            SetStatus("确认失败，目标对话框控件不可用。");
            return;
        }

        SetStatus($"已提交：{candidate.DisplayName}");
    }

    private void SearchModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        SwitchMode(OverlayMode.Search);
    }

    private void RecentModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        SwitchMode(OverlayMode.Recent);
    }

    private void TrayModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        SwitchMode(OverlayMode.Tray);
    }

    private void ExplorerModeButton_OnClick(object sender, RoutedEventArgs e)
    {
        SwitchMode(OverlayMode.Explorer);
    }

    private void SortButton_OnClick(object sender, RoutedEventArgs e)
    {
        _sortIndex = (_sortIndex + 1) % _sortOptions.Count;
        UpdateSortButtonLabel();

        if (_currentMode is OverlayMode.Search or OverlayMode.Recent)
        {
            _searchDebounceTimer.Stop();
            _ = RefreshCandidatesAsync(immediate: true);
        }
    }

    private void SwitchMode(OverlayMode mode)
    {
        if (_currentMode == mode)
        {
            return;
        }

        _currentMode = mode;
        ApplyModeUiState();

        _searchDebounceTimer.Stop();
        _ = RefreshCandidatesAsync(immediate: true);
    }

    private void ApplyModeUiState()
    {
        SearchModeButton.IsChecked = _currentMode == OverlayMode.Search;
        RecentModeButton.IsChecked = _currentMode == OverlayMode.Recent;
        TrayModeButton.IsChecked = _currentMode == OverlayMode.Tray;
        ExplorerModeButton.IsChecked = _currentMode == OverlayMode.Explorer;
        TypeFilterButton.IsEnabled = !_isFolderDialog;
        TypeFilterButton.IsChecked = IsTypeFilterActive;
        TypeFilterButton.ToolTip = _isFolderDialog
            ? "文件夹选择时不启用类型过滤"
            : "按文件类型过滤";
        SortButton.IsEnabled = _currentMode is OverlayMode.Search or OverlayMode.Recent;
    }

    private void UpdateSortButtonLabel()
    {
        SortLabelText.Text = _sortOptions[_sortIndex].Label;
    }

    private void TypeFilterButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isFolderDialog)
        {
            ApplyModeUiState();
            return;
        }

        _preferredTypeFilterEnabled = !_preferredTypeFilterEnabled;
        _preferencesRepository.SetTypeFilterEnabled(_preferredTypeFilterEnabled);
        ApplyModeUiState();

        _searchDebounceTimer.Stop();
        _ = RefreshCandidatesAsync(immediate: true);
    }

    private async void Window_OnPreviewDrop(object sender, System.Windows.DragEventArgs e)
    {
        var paths = ExtractDroppedPaths(e.Data);
        if (paths.Count == 0)
        {
            SetStatus("未识别到可加入托盘的文件。");
            return;
        }

        var added = _trayRepository.AddMany(paths);
        SetStatus(added > 0 ? $"已加入托盘：{added} 项。" : "拖入项已在托盘中。");

        if (_currentMode == OverlayMode.Tray)
        {
            await RefreshCandidatesAsync(immediate: true);
        }

        e.Handled = true;
    }

    private void Window_OnPreviewDragOver(object sender, System.Windows.DragEventArgs e)
    {
        e.Effects = HasDroppedFiles(e.Data)
            ? System.Windows.DragDropEffects.Copy
            : System.Windows.DragDropEffects.None;
        e.Handled = true;
    }

    private static bool HasDroppedFiles(System.Windows.IDataObject data)
    {
        return data.GetDataPresent(System.Windows.DataFormats.FileDrop)
            || data.GetDataPresent("FileNameW")
            || data.GetDataPresent("FileName");
    }

    private static IReadOnlyList<string> ExtractDroppedPaths(System.Windows.IDataObject data)
    {
        if (data.GetDataPresent(System.Windows.DataFormats.FileDrop) &&
            data.GetData(System.Windows.DataFormats.FileDrop) is string[] standardPaths &&
            standardPaths.Length > 0)
        {
            return standardPaths;
        }

        foreach (var format in new[] { "FileNameW", "FileName" })
        {
            if (!data.GetDataPresent(format))
            {
                continue;
            }

            var raw = data.GetData(format);
            if (raw is string single && !string.IsNullOrWhiteSpace(single))
            {
                return [single];
            }

            if (raw is string[] many && many.Length > 0)
            {
                return many;
            }
        }

        return [];
    }

    private static int DipToPx(double dip, double scale)
    {
        return (int)Math.Round(dip * scale);
    }

    private static double PxToDip(int px, double scale)
    {
        return px / scale;
    }

    private static NativeMethods.RECT GetWorkArea(IntPtr monitor, NativeMethods.RECT fallback)
    {
        if (monitor == IntPtr.Zero)
        {
            return fallback;
        }

        var info = new NativeMethods.MONITORINFO
        {
            cbSize = System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.MONITORINFO>(),
        };

        return NativeMethods.GetMonitorInfo(monitor, ref info)
            ? info.rcWork
            : fallback;
    }

    private static bool HasAllowedExtension(string fullPath, IReadOnlySet<string> allowedExtensions)
    {
        var extension = Path.GetExtension(fullPath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            return false;
        }

        return allowedExtensions.Contains(extension.ToLowerInvariant());
    }

    private bool AreItemsEquivalent(IReadOnlyList<FileCandidate> items)
    {
        if (_items.Count != items.Count)
        {
            return false;
        }

        for (var i = 0; i < items.Count; i++)
        {
            var left = _items[i];
            var right = items[i];
            if (!string.Equals(left.FullPath, right.FullPath, StringComparison.OrdinalIgnoreCase) ||
                left.IsDirectory != right.IsDirectory ||
                left.Source != right.Source)
            {
                return false;
            }
        }

        return true;
    }

    private static T? FindDescendant<T>(DependencyObject root) where T : DependencyObject
    {
        var childCount = VisualTreeHelper.GetChildrenCount(root);
        for (var i = 0; i < childCount; i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is T typed)
            {
                return typed;
            }

            var nested = FindDescendant<T>(child);
            if (nested is not null)
            {
                return nested;
            }
        }

        return null;
    }
}
