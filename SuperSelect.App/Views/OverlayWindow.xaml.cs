using System.Collections.ObjectModel;
using System.ComponentModel;
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

    private const double UltraCompactHeightDip = 66.0;
    private const double CompactHeightDip = 108.0;
    private const double ExpandedHeightDip = 350.0;
    private const double MinUsableHeightDip = 66.0;
    private const double HeaderAndStatusHeightDip = 112.0;
    private const int OverlayGapPx = 0;
    private const int InitialQueryResultLimit = 320;
    private const int ExpandedQueryResultLimit = 1200;
    private const int ExpandedTypeFilterQueryResultLimit = 5000;
    private const int DisplayResultLimit = 400;
    private const int TypeFilterScanPageSize = 2000;
    private const int AutomationActionTimeoutMs = 900;

    private readonly EverythingService _everythingService;
    private readonly TrayRepository _trayRepository;
    private readonly ExplorerWindowService _explorerWindowService;
    private readonly UserPreferencesRepository _preferencesRepository;
    private readonly ObservableCollection<FileCandidate> _items = [];
    private readonly DispatcherTimer _searchDebounceTimer;
    private readonly DispatcherTimer _dialogMonitorTimer;
    private readonly List<OptionItem<EverythingSortOption>> _sortOptions;
    private readonly SemaphoreSlim _automationOperationGate = new(1, 1);

    private CancellationTokenSource? _refreshCts;
    private CancellationTokenSource? _singleSelectCts;
    private FileDialogAutomationController? _dialogController;
    private ScrollViewer? _resultListScrollViewer;
    private IntPtr _dialogHwnd;
    private int _wheelDeltaAccumulator;
    private OverlayMode _currentMode = OverlayMode.Search;
    private int _sortIndex;
    private bool _preferredTypeFilterEnabled;
    private bool _isFolderDialog;
    private string _lastFileTypeFilterText = string.Empty;
    private int _dialogMonitorInFlight;
    private long _singleSelectVersion;
    private bool _allowClose;

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
        Interlocked.Increment(ref _singleSelectVersion);
        
        Task.Run(() =>
        {
            var isFolder = _dialogController.IsFolderSelectionDialog();
            var filter = _dialogController.GetCurrentFileTypeFilterText();
            _ = Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_dialogHwnd != dialogHwnd) return;
                _isFolderDialog = isFolder;
                _lastFileTypeFilterText = filter;
                ApplyModeUiState();
                _ = RefreshCandidatesAsync(immediate: true);
            }));
        });

        ApplyModeUiState();

        if (!IsVisible)
        {
            var overlayHwnd = new WindowInteropHelper(this).EnsureHandle();
            NativeMethods.SetWindowLongPtrEx(overlayHwnd, NativeMethods.GWL_HWNDPARENT, _dialogHwnd);
            Show();
        }
        else
        {
            var overlayHwnd = new WindowInteropHelper(this).Handle;
            NativeMethods.SetWindowLongPtrEx(overlayHwnd, NativeMethods.GWL_HWNDPARENT, _dialogHwnd);
        }

        _dialogMonitorTimer.Start();
        PositionToDialog();
    }

    public void DetachDialog()
    {
        _dialogHwnd = IntPtr.Zero;
        _dialogController = null;
        _wheelDeltaAccumulator = 0;
        _dialogMonitorTimer.Stop();
        Interlocked.Exchange(ref _dialogMonitorInFlight, 0);
        _isFolderDialog = false;
        _lastFileTypeFilterText = string.Empty;
        ApplyModeUiState();
        Interlocked.Increment(ref _singleSelectVersion);

        CancelAndDisposeRefreshCts();
        CancelAndDisposeSingleSelectCts();

        _items.Clear();
        SearchBox.Text = string.Empty;
        ResultHost.Visibility = Visibility.Collapsed;
        SetStatus(null);
        Height = UltraCompactHeightDip;

        if (IsVisible)
        {
            var overlayHwnd = new WindowInteropHelper(this).Handle;
            if (overlayHwnd != IntPtr.Zero)
            {
                NativeMethods.SetWindowLongPtrEx(overlayHwnd, NativeMethods.GWL_HWNDPARENT, IntPtr.Zero);
            }
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

    public void Shutdown()
    {
        _allowClose = true;
        CancelAndDisposeRefreshCts();
        CancelAndDisposeSingleSelectCts();
        Close();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_allowClose)
        {
            base.OnClosing(e);
            return;
        }

        e.Cancel = true;
        DetachDialog();
    }

    private async Task RefreshCandidatesAsync(bool immediate = false)
    {
        if (_dialogController is null)
        {
            return;
        }

        var token = ReplaceRefreshToken();

        var mode = _currentMode;
        var keyword = SearchBox.Text.Trim();
        var sort = CurrentSort;

        if (mode == OverlayMode.Search && string.IsNullOrWhiteSpace(keyword))
        {
            SetStatus(null);
            UpdateItems([]);
            return;
        }

        if (!immediate)
        {
            try
            {
                await Task.Delay(15, token);
            }
            catch (OperationCanceledException)
            {
                return;
            }
        }

        var initialOutcome = await QueryCandidatesByModeAsync(mode, keyword, sort, InitialQueryResultLimit, token);
        if (initialOutcome is null || token.IsCancellationRequested)
        {
            return;
        }

        var (initialFilteredItems, initialStatus) = ApplyTypeFilter(initialOutcome.Value.Items, initialOutcome.Value.StatusMessage, initialOutcome.Value.CanApplyTypeFilter, initialOutcome.Value.Caption);
        var (initialDisplayItems, initialTrimmed) = CapDisplayItems(initialFilteredItems);
        SetStatus(BuildStatus(initialStatus, initialTrimmed, expandedScan: false));
        UpdateItems(initialDisplayItems);

        var expandedLimit = IsTypeFilterActive ? ExpandedTypeFilterQueryResultLimit : ExpandedQueryResultLimit;
        var shouldExpand =
            mode is OverlayMode.Search or OverlayMode.Recent &&
            expandedLimit > InitialQueryResultLimit &&
            initialOutcome.Value.Items.Count >= InitialQueryResultLimit &&
            (IsTypeFilterActive || initialDisplayItems.Count < DisplayResultLimit);

        if (!shouldExpand)
        {
            return;
        }

        if (IsTypeFilterActive &&
            initialOutcome.Value.CanApplyTypeFilter &&
            mode is OverlayMode.Search or OverlayMode.Recent &&
            _dialogController is not null)
        {
            var allowedExtensions = _dialogController.GetAllowedFileExtensions(_lastFileTypeFilterText);
            if (allowedExtensions.Count > 0)
            {
                var pagedOutcome = await QueryTypeFilteredCandidatesByPagingAsync(
                    mode,
                    keyword,
                    sort,
                    initialOutcome.Value.Caption,
                    allowedExtensions,
                    token);

                if (pagedOutcome is not null && !token.IsCancellationRequested)
                {
                    SetStatus(pagedOutcome.Value.StatusMessage);
                    UpdateItems(pagedOutcome.Value.Items);
                }

                return;
            }
        }

        var expandedOutcome = await QueryCandidatesByModeAsync(mode, keyword, sort, expandedLimit, token);
        if (expandedOutcome is null || token.IsCancellationRequested)
        {
            return;
        }

        var (expandedFilteredItems, expandedStatus) = ApplyTypeFilter(expandedOutcome.Value.Items, expandedOutcome.Value.StatusMessage, expandedOutcome.Value.CanApplyTypeFilter, expandedOutcome.Value.Caption);
        var (expandedDisplayItems, expandedTrimmed) = CapDisplayItems(expandedFilteredItems);
        SetStatus(BuildStatus(expandedStatus, expandedTrimmed, expandedScan: true));
        UpdateItems(expandedDisplayItems);
    }

    private async Task<(IReadOnlyList<FileCandidate> Items, string StatusMessage)?> QueryTypeFilteredCandidatesByPagingAsync(
        OverlayMode mode,
        string keyword,
        EverythingSortOption sort,
        string caption,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken token)
    {
        var result = new List<FileCandidate>(DisplayResultLimit);
        var offset = 0;
        var exhausted = false;
        var scannedPages = 0;

        while (!token.IsCancellationRequested)
        {
            var pageResult = await QueryEverythingTypeFilteredPageByModeAsync(
                mode,
                keyword,
                sort,
                offset,
                TypeFilterScanPageSize,
                allowedExtensions,
                token);

            if (pageResult.RawCount == 0)
            {
                exhausted = true;
                break;
            }

            scannedPages++;
            foreach (var candidate in pageResult.Items)
            {
                result.Add(candidate);
                if (result.Count >= DisplayResultLimit)
                {
                    var cappedStatus = $"{caption}结果：至少 {DisplayResultLimit} 条（类型过滤，已分段扫描 {scannedPages} 页）";
                    return (result, cappedStatus);
                }
            }

            if (pageResult.RawCount < TypeFilterScanPageSize)
            {
                exhausted = true;
                break;
            }

            offset += pageResult.RawCount;
        }

        if (token.IsCancellationRequested)
        {
            return null;
        }

        var status = exhausted
            ? $"{caption}结果：{result.Count} 条（类型过滤，已全量分段扫描 {scannedPages} 页）"
            : $"{caption}结果：{result.Count} 条（类型过滤）";
        return (result, status);
    }

    private Task<(IReadOnlyList<FileCandidate> Items, int RawCount)> QueryEverythingTypeFilteredPageByModeAsync(
        OverlayMode mode,
        string keyword,
        EverythingSortOption sort,
        int offset,
        int pageSize,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken token)
    {
        return mode switch
        {
            OverlayMode.Search when !_isFolderDialog => _everythingService.SearchFilesFilteredPageAsync(keyword, sort, offset, pageSize, allowedExtensions, token),
            OverlayMode.Recent when !_isFolderDialog => _everythingService.RecentFilesFilteredPageAsync(sort, offset, pageSize, allowedExtensions, token),
            _ => Task.FromResult((Items: (IReadOnlyList<FileCandidate>)[], RawCount: 0)),
        };
    }

    private async Task<(IReadOnlyList<FileCandidate> Items, string? StatusMessage, bool CanApplyTypeFilter, string Caption)?> QueryCandidatesByModeAsync(
        OverlayMode mode,
        string keyword,
        EverythingSortOption sort,
        int resultLimit,
        CancellationToken token)
    {
        IReadOnlyList<FileCandidate> items = [];
        string? statusMessage = string.Empty;
        var canApplyTypeFilter = true;
        var caption = "结果";

        try
        {
            switch (mode)
            {
                case OverlayMode.Search:
                    items = _isFolderDialog
                        ? await _everythingService.SearchFoldersAsync(keyword, sort, resultLimit, token)
                        : await _everythingService.SearchFilesAsync(keyword, sort, resultLimit, token);

                    caption = _isFolderDialog ? "文件夹搜索" : "搜索";
                    if (!_everythingService.IsAvailable)
                    {
                        statusMessage = $"Everything 不可用：{_everythingService.LastErrorMessage}";
                        canApplyTypeFilter = false;
                    }
                    else
                    {
                        statusMessage = items.Count == 0
                            ? (_isFolderDialog ? "没有文件夹搜索结果。" : "没有搜索结果。")
                            : $"{caption}结果：{items.Count} 条。";
                    }

                    break;
                case OverlayMode.Recent:
                    items = _isFolderDialog
                        ? await _everythingService.RecentFoldersAsync(sort, resultLimit, token)
                        : await _everythingService.RecentFilesAsync(sort, resultLimit, token);

                    caption = _isFolderDialog ? "最近文件夹" : "最近文件";
                    if (!_everythingService.IsAvailable)
                    {
                        statusMessage = $"Everything 不可用：{_everythingService.LastErrorMessage}";
                        canApplyTypeFilter = false;
                    }
                    else
                    {
                        statusMessage = items.Count == 0
                            ? (_isFolderDialog ? "没有最近文件夹结果。" : "没有最近文件结果。")
                            : $"{caption}：{items.Count} 条。";
                    }

                    break;
                case OverlayMode.Tray:
                    items = _trayRepository.Query(string.Empty);
                    if (_isFolderDialog)
                    {
                        items = items.Where(candidate => candidate.IsDirectory).ToList();
                    }

                    caption = _isFolderDialog ? "托盘文件夹" : "托盘文件";
                    statusMessage = items.Count == 0 ? null : $"{caption}：{items.Count} 条。";
                    break;
                case OverlayMode.Explorer:
                    items = _explorerWindowService.GetOpenLocations(string.Empty);
                    caption = "资源管理器路径";
                    statusMessage = items.Count == 0
                        ? "未检测到可用资源管理器路径。"
                        : $"{caption}：{items.Count} 条。";
                    canApplyTypeFilter = false;
                    break;
            }
        }
        catch (OperationCanceledException)
        {
            return null;
        }

        return (items, statusMessage, canApplyTypeFilter, caption);
    }

    private (IReadOnlyList<FileCandidate> Items, string? StatusMessage) ApplyTypeFilter(
        IReadOnlyList<FileCandidate> items,
        string? statusMessage,
        bool canApplyDialogTypeFilter,
        string caption)
    {
        if (_isFolderDialog)
        {
            canApplyDialogTypeFilter = false;
        }

        if (IsTypeFilterActive && canApplyDialogTypeFilter && _dialogController is not null)
        {
            var allowedExtensions = _dialogController.GetAllowedFileExtensions(_lastFileTypeFilterText);
            if (allowedExtensions.Count > 0)
            {
                items = items
                    .Where(candidate => candidate.IsDirectory || HasAllowedExtension(candidate.FullPath, allowedExtensions))
                    .ToList();
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

        return (items, statusMessage);
    }

    private static (IReadOnlyList<FileCandidate> Items, bool Trimmed) CapDisplayItems(IReadOnlyList<FileCandidate> items)
    {
        if (items.Count <= DisplayResultLimit)
        {
            return (items, false);
        }

        return (items.Take(DisplayResultLimit).ToList(), true);
    }

    private static string? BuildStatus(string? statusMessage, bool trimmed, bool expandedScan)
    {
        if (!trimmed && !expandedScan)
        {
            return string.IsNullOrWhiteSpace(statusMessage) ? null : statusMessage;
        }

        var baseText = string.IsNullOrWhiteSpace(statusMessage) ? "查询结果" : statusMessage!;
        if (trimmed && expandedScan)
        {
            return $"{baseText}（显示前 {DisplayResultLimit} 条，已扩展扫描）";
        }

        if (trimmed)
        {
            return $"{baseText}（显示前 {DisplayResultLimit} 条）";
        }

        return $"{baseText}（已扩展扫描）";
    }

    private CancellationToken ReplaceRefreshToken()
    {
        var next = new CancellationTokenSource();
        var previous = Interlocked.Exchange(ref _refreshCts, next);
        if (previous is not null)
        {
            previous.Cancel();
            previous.Dispose();
        }

        return next.Token;
    }

    private void CancelAndDisposeRefreshCts()
    {
        var previous = Interlocked.Exchange(ref _refreshCts, null);
        if (previous is null)
        {
            return;
        }

        previous.Cancel();
        previous.Dispose();
    }

    private CancellationToken ReplaceSingleSelectToken()
    {
        var next = new CancellationTokenSource();
        var previous = Interlocked.Exchange(ref _singleSelectCts, next);
        if (previous is not null)
        {
            previous.Cancel();
            previous.Dispose();
        }

        return next.Token;
    }

    private void CancelAndDisposeSingleSelectCts()
    {
        var previous = Interlocked.Exchange(ref _singleSelectCts, null);
        if (previous is null)
        {
            return;
        }

        previous.Cancel();
        previous.Dispose();
    }

    private void UpdateItems(IReadOnlyList<FileCandidate> items)
    {
        var previousVisibility = ResultHost.Visibility;
        var previousHeight = Height;
        var changed = !AreItemsEquivalent(items);

        if (changed)
        {
            var commonPrefixLength = 0;
            var minCount = Math.Min(_items.Count, items.Count);
            while (commonPrefixLength < minCount &&
                   CandidateEquivalent(_items[commonPrefixLength], items[commonPrefixLength]))
            {
                commonPrefixLength++;
            }

            for (var i = _items.Count - 1; i >= commonPrefixLength; i--)
            {
                _items.RemoveAt(i);
            }

            for (var i = commonPrefixLength; i < items.Count; i++)
            {
                _items.Add(items[i]);
            }

            if (_items.Count > 0)
            {
                ResultList.SelectedIndex = 0;
            }
            else
            {
                ResultList.SelectedItem = null;
            }
        }
        else if (_items.Count > 0 && ResultList.SelectedIndex < 0)
        {
            ResultList.SelectedIndex = 0;
        }

        ResultHost.Visibility = _items.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        var statusVisible = StatusText.Visibility == Visibility.Visible;
        var baseCompactHeight = statusVisible ? CompactHeightDip : UltraCompactHeightDip;
        Height = _items.Count > 0 ? ExpandedHeightDip : baseCompactHeight;

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

    private void SetStatus(string? message)
    {
        if (string.IsNullOrEmpty(message))
        {
            StatusText.Visibility = Visibility.Collapsed;
            StatusText.Text = string.Empty;
        }
        else
        {
            StatusText.Visibility = Visibility.Visible;
            StatusText.Text = message;
        }
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
        if (Interlocked.CompareExchange(ref _dialogMonitorInFlight, 1, 0) != 0)
        {
            return;
        }

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

            if (!NativeMethods.IsWindowVisible(_dialogHwnd))
            {
                DetachDialog();
                return;
            }

            var controller = _dialogController;
            var isTypeFilterActive = IsTypeFilterActive;
            var currentMode = _currentMode;

            var folderModeTask = Task.Run(() => controller.IsFolderSelectionDialog());
            var filterTextTask = Task.Run(() => 
            {
                if (isTypeFilterActive && currentMode != OverlayMode.Explorer)
                {
                    return controller.GetCurrentFileTypeFilterText();
                }
                return _lastFileTypeFilterText;
            });

            var monitorTask = Task.WhenAll(folderModeTask, filterTextTask);
            try
            {
                await monitorTask.WaitAsync(TimeSpan.FromMilliseconds(320));
            }
            catch (TimeoutException)
            {
                return;
            }

            if (_dialogController != controller) return;

            var folderMode = folderModeTask.Result;
            var folderModeChanged = folderMode != _isFolderDialog;
            if (folderModeChanged)
            {
                _isFolderDialog = folderMode;
                ApplyModeUiState();
            }

            var filterChanged = false;
            if (isTypeFilterActive && currentMode != OverlayMode.Explorer)
            {
                var filterText = filterTextTask.Result;
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
        finally
        {
            Interlocked.Exchange(ref _dialogMonitorInFlight, 0);
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

    private void SearchBox_OnPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Down)
        {
            MoveSelection(+1);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Up)
        {
            MoveSelection(-1);
            e.Handled = true;
            return;
        }

        if (e.Key != Key.Enter)
        {
            return;
        }

        var candidate = GetCurrentOrFirstCandidate();
        if (candidate is null)
        {
            return;
        }

        _ = ExecuteConfirmAsync(candidate);
        e.Handled = true;
    }

    private void ResultList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ResultList.SelectedItem is FileCandidate candidate)
        {
            _ = ExecuteConfirmAsync(candidate);
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
        if (e.Key == Key.Down)
        {
            MoveSelection(+1);
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Up)
        {
            MoveSelection(-1);
            e.Handled = true;
            return;
        }

        if (e.Key != Key.Enter)
        {
            return;
        }

        var candidate = GetCurrentOrFirstCandidate();
        if (candidate is null)
        {
            return;
        }

        _ = ExecuteConfirmAsync(candidate);
        e.Handled = true;
    }

    private async Task ExecuteSingleAsync(FileCandidate candidate)
    {
        if (_dialogController is null)
        {
            return;
        }

        var token = ReplaceSingleSelectToken();
        var operationVersion = Interlocked.Increment(ref _singleSelectVersion);
        var controller = _dialogController;
        var gateAcquired = false;
        try
        {
            await _automationOperationGate.WaitAsync(token);
            gateAcquired = true;
        }
        catch (OperationCanceledException)
        {
            return;
        }

        bool ok;
        try
        {
            var automationTask = Task.Run(
                () => controller.TryPrimeSelectionAsync(candidate, token).GetAwaiter().GetResult(),
                token);

            ok = await WaitForAutomationResultAsync(
                automationTask,
                "OverlayWindow.ExecuteSingleAsync.TryPrimeSelection",
                token);
        }
        catch (OperationCanceledException)
        {
            return;
        }
        finally
        {
            if (gateAcquired)
            {
                _automationOperationGate.Release();
            }
        }

        if (token.IsCancellationRequested)
        {
            return;
        }

        if (operationVersion != Interlocked.Read(ref _singleSelectVersion) || _dialogController != controller)
        {
            return;
        }

        if (!ok)
        {
            SetStatus("单击定位失败或超时，目标对话框可能繁忙。");
            return;
        }

        SetStatus($"已定位：{candidate.DisplayName}");
    }

    private async Task ExecuteConfirmAsync(FileCandidate candidate)
    {
        if (_dialogController is null)
        {
            return;
        }

        Interlocked.Increment(ref _singleSelectVersion);
        CancelAndDisposeSingleSelectCts();

        var controller = _dialogController;
        await _automationOperationGate.WaitAsync();
        bool ok;
        try
        {
            var automationTask = Task.Run(
                () => controller.TryConfirmSelection(candidate));

            ok = await WaitForAutomationResultAsync(
                automationTask,
                "OverlayWindow.ExecuteConfirmAsync.TryConfirmSelection",
                CancellationToken.None);
        }
        finally
        {
            _automationOperationGate.Release();
        }

        if (_dialogController != controller)
        {
            return;
        }

        if (!ok)
        {
            SetStatus("确认失败或超时，目标对话框可能繁忙。");
            return;
        }

        SetStatus($"已提交：{candidate.DisplayName}");
    }

    private static async Task<bool> WaitForAutomationResultAsync(
        Task<bool> task,
        string context,
        CancellationToken cancellationToken)
    {
        try
        {
            return await task.WaitAsync(TimeSpan.FromMilliseconds(AutomationActionTimeoutMs), cancellationToken);
        }
        catch (TimeoutException)
        {
            _ = task.ContinueWith(
                t =>
                {
                    if (t.Exception is not null)
                    {
                        AppLogger.LogException($"{context}.LateFault", t.Exception);
                    }
                },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted,
                TaskScheduler.Default);

            AppLogger.LogWarning($"{context} timeout after {AutomationActionTimeoutMs}ms.");
            return false;
        }
        catch (OperationCanceledException)
        {
            return false;
        }
        catch (Exception ex)
        {
            AppLogger.LogException(context, ex, throttle: TimeSpan.FromSeconds(2));
            return false;
        }
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

    private void ResultList_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (_currentMode != OverlayMode.Tray)
        {
            e.Handled = true;
        }
    }

    private void ListBoxItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_currentMode == OverlayMode.Tray && sender is ListBoxItem { ContextMenu: { } menu } item)
        {
            menu.PlacementTarget = item;
            menu.IsOpen = true;
            e.Handled = true;
        }
    }

    private void ListBoxItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount != 1 || sender is not ListBoxItem { DataContext: FileCandidate candidate })
        {
            return;
        }

        if (!ReferenceEquals(ResultList.SelectedItem, candidate))
        {
            ResultList.SelectedItem = candidate;
        }

        _ = ExecuteSingleAsync(candidate);
    }

    private void TrayMenuItem_Pin_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem && 
            menuItem.DataContext is FileCandidate candidate)
        {
            _trayRepository.PinToTop(candidate.FullPath);
            _ = RefreshCandidatesAsync(immediate: true);
        }
    }

    private void TrayMenuItem_Delete_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem && 
            menuItem.DataContext is FileCandidate candidate)
        {
            _trayRepository.Remove(candidate.FullPath);
            _ = RefreshCandidatesAsync(immediate: true);
        }
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

    private FileCandidate? GetCurrentOrFirstCandidate()
    {
        return ResultList.SelectedItem as FileCandidate ?? _items.FirstOrDefault();
    }

    private void MoveSelection(int delta)
    {
        if (_items.Count == 0)
        {
            return;
        }

        var currentIndex = ResultList.SelectedIndex;
        int nextIndex;

        if (currentIndex < 0)
        {
            nextIndex = delta >= 0 ? 0 : _items.Count - 1;
        }
        else
        {
            nextIndex = Math.Clamp(currentIndex + delta, 0, _items.Count - 1);
        }

        if (nextIndex == currentIndex)
        {
            return;
        }

        ResultList.SelectedIndex = nextIndex;
        if (ResultList.SelectedItem is FileCandidate selected)
        {
            ResultList.ScrollIntoView(selected);
        }
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
            if (!CandidateEquivalent(left, right))
            {
                return false;
            }
        }

        return true;
    }

    private static bool CandidateEquivalent(FileCandidate left, FileCandidate right)
    {
        return string.Equals(left.FullPath, right.FullPath, StringComparison.OrdinalIgnoreCase) &&
               left.IsDirectory == right.IsDirectory &&
               left.Source == right.Source;
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
