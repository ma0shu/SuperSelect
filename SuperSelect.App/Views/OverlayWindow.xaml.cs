using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
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

    private const double UltraCompactHeightDip = 48.0;
    private const double CompactHeightDip = 76.0;
    private const double ExpandedHeightDip = 216.0;
    private const double MinUsableHeightDip = 48.0;
    private const double HeaderAndStatusHeightDip = 72.0;
    private const double ResultSummaryHeightDip = 14.0;
    private const int OverlayGapPx = 0;
    private const int InitialQueryResultLimit = 320;
    private const int ExpandedQueryResultLimit = 1200;
    private const int ExpandedTypeFilterQueryResultLimit = 5000;
    private const int DisplayResultLimit = 400;
    private const int TypeFilterScanPageSize = 2000;
    private const int AutomationActionTimeoutMs = 900;
    private const int SearchDebounceIntervalMs = 80;
    private static readonly TimeSpan DialogMonitorLivenessInterval = TimeSpan.FromMilliseconds(140);
    private static readonly TimeSpan DialogMonitorBusyInterval = TimeSpan.FromMilliseconds(380);
    private static readonly TimeSpan DialogMonitorIdleInterval = TimeSpan.FromMilliseconds(760);
    private static readonly TimeSpan DialogMonitorTimeout = TimeSpan.FromMilliseconds(260);
    private static readonly TimeSpan FolderStateProbeInterval = TimeSpan.FromMilliseconds(1200);

    private readonly EverythingService _everythingService;
    private readonly TrayRepository _trayRepository;
    private readonly ExplorerWindowService _explorerWindowService;
    private readonly UserPreferencesRepository _preferencesRepository;
    private readonly ObservableCollection<FileCandidate> _items = [];
    private readonly DispatcherTimer _searchDebounceTimer;
    private readonly DispatcherTimer _dialogMonitorTimer;
    private readonly DispatcherTimer _positionUpdateTimer;
    private readonly DispatcherTimer _dialogRectPollTimer;
    private readonly List<OptionItem<EverythingSortOption>> _sortOptions;
    private readonly ContextMenu _sortContextMenu;
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
    private DateTime _lastDialogSnapshotProbeUtc = DateTime.MinValue;
    private DateTime _lastFolderStateProbeUtc = DateTime.MinValue;
    private long _singleSelectVersion;
    private long _refreshRequestVersion;
    private long _refreshProcessedVersion;
    private int _refreshLoopInFlight;
    private bool _refreshImmediateRequested;
    private readonly object _refreshRequestSync = new();
    private bool _allowClose;
    private bool _dropInteropConfigured;
    private bool _positionUpdateRequested;
    private bool _positionRecomputeRequested;
    private bool _hasAppliedPosition;
    private int _lastOverlayLeftPx;
    private int _lastOverlayTopPx;
    private int _lastOverlayWidthPx;
    private int _lastOverlayHeightPx;
    private IntPtr _overlayHwnd;
    private bool _hasLastDialogWindowRect;
    private NativeMethods.RECT _lastDialogWindowRect;
    private bool _hasLastDialogVisualRect;
    private NativeMethods.RECT _lastDialogVisualRect;
    private Visibility _lastMeasuredStatusVisibility = Visibility.Collapsed;
    private Visibility _lastMeasuredSummaryVisibility = Visibility.Collapsed;

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
        _sortContextMenu = CreateSortContextMenu();

        _searchDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(SearchDebounceIntervalMs),
        };
        _searchDebounceTimer.Tick += SearchDebounceTimer_OnTick;
        _dialogMonitorTimer = new DispatcherTimer();
        _dialogMonitorTimer.Tick += DialogMonitorTimer_OnTick;
        _positionUpdateTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(16)
        };
        _positionUpdateTimer.Tick += PositionUpdateTimer_OnTick;
        _dialogRectPollTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(16)
        };
        _dialogRectPollTimer.Tick += DialogRectPollTimer_OnTick;

        ResultList.ItemsSource = _items;
        SortButton.ContextMenu = _sortContextMenu;
        ApplyModeUiState();
        UpdateSortButtonLabel();

        ShowActivated = false;
    }

    private void PositionUpdateTimer_OnTick(object? sender, EventArgs e)
    {
        if (!_positionUpdateRequested)
        {
            _positionUpdateTimer.Stop();
            return;
        }

        _positionUpdateRequested = false;
        _positionUpdateTimer.Stop();
        PositionToDialog();
    }

    private void RequestPositionUpdate()
    {
        if (!IsVisible || _dialogHwnd == IntPtr.Zero)
        {
            return;
        }

        _positionUpdateRequested = true;
        _positionRecomputeRequested = true;
        if (!_positionUpdateTimer.IsEnabled)
        {
            _positionUpdateTimer.Start();
        }
    }

    private void DialogRectPollTimer_OnTick(object? sender, EventArgs e)
    {
        if (!IsVisible || _dialogHwnd == IntPtr.Zero)
        {
            _dialogRectPollTimer.Stop();
            return;
        }

        if (!NativeMethods.IsWindow(_dialogHwnd) || !NativeMethods.IsWindowVisible(_dialogHwnd))
        {
            DetachDialog();
            return;
        }

        if (!NativeMethods.GetWindowRect(_dialogHwnd, out var dialogRect))
        {
            return;
        }

        if (!_positionRecomputeRequested &&
            _hasLastDialogWindowRect &&
            RectEquals(_lastDialogWindowRect, dialogRect))
        {
            return;
        }

        PositionToDialog(dialogRect);
    }

    public void AttachToDialog(IntPtr dialogHwnd)
    {
        _dialogHwnd = dialogHwnd;
        _dialogController = new FileDialogAutomationController(dialogHwnd);
        Interlocked.Increment(ref _singleSelectVersion);
        CancelAndDisposeRefreshCts();
        _searchDebounceTimer.Stop();
        _explorerWindowService.EnsureSnapshotFreshAsync(force: false);
        ResetModeForDialog();

        var controller = _dialogController;
        if (controller is not null)
        {
            _ = Task.Run(() =>
            {
                var snapshot = CaptureDialogSnapshot(
                    controller,
                    includeFilterText: true,
                    fallbackFilterText: string.Empty,
                    includeFolderState: true,
                    fallbackFolderState: _isFolderDialog);
                _ = Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_dialogHwnd != dialogHwnd || _dialogController != controller)
                    {
                        return;
                    }

                    _isFolderDialog = snapshot.IsFolderDialog;
                    _lastFileTypeFilterText = snapshot.FilterText;
                    _lastFolderStateProbeUtc = DateTime.UtcNow;
                    if (_currentMode == OverlayMode.Search &&
                        _explorerWindowService.HasOpenLocationsCached() &&
                        string.IsNullOrWhiteSpace(SearchBox.Text))
                    {
                        _currentMode = OverlayMode.Explorer;
                    }
                    ApplyModeUiState();
                    RequestCandidatesRefresh(immediate: true);
                }));
            });
        }

        ApplyModeUiState();
        _overlayHwnd = new WindowInteropHelper(this).EnsureHandle();
        ConfigureDropInteropCompatibility(_overlayHwnd);
        NativeMethods.SetWindowLongPtrEx(_overlayHwnd, NativeMethods.GWL_HWNDPARENT, _dialogHwnd);
        if (!IsVisible)
        {
            Show();
        }

        _hasLastDialogWindowRect = false;
        _hasLastDialogVisualRect = false;
        _positionRecomputeRequested = true;
        _lastDialogSnapshotProbeUtc = DateTime.MinValue;
        _dialogMonitorTimer.Interval = DialogMonitorLivenessInterval;
        _dialogMonitorTimer.Start();
        if (!_dialogRectPollTimer.IsEnabled)
        {
            _dialogRectPollTimer.Start();
        }
        PositionToDialog();
        RequestCandidatesRefresh(immediate: true);
    }

    public void DetachDialog()
    {
        _dialogHwnd = IntPtr.Zero;
        _dialogController = null;
        _wheelDeltaAccumulator = 0;
        _dialogMonitorTimer.Stop();
        _dialogRectPollTimer.Stop();
        _positionUpdateTimer.Stop();
        _positionUpdateRequested = false;
        _positionRecomputeRequested = false;
        Interlocked.Exchange(ref _dialogMonitorInFlight, 0);
        _lastDialogSnapshotProbeUtc = DateTime.MinValue;
        _lastFolderStateProbeUtc = DateTime.MinValue;
        _isFolderDialog = false;
        _lastFileTypeFilterText = string.Empty;
        ApplyModeUiState();
        Interlocked.Increment(ref _singleSelectVersion);
        lock (_refreshRequestSync) _refreshImmediateRequested = false;

        CancelAndDisposeRefreshCts();
        CancelAndDisposeSingleSelectCts();

        _items.Clear();
        SearchBox.Text = string.Empty;
        ResultHost.Visibility = Visibility.Collapsed;
        SetStatus(null);
        SetResultSummary(null);
        Height = UltraCompactHeightDip;
        _hasAppliedPosition = false;
        _lastOverlayLeftPx = 0;
        _lastOverlayTopPx = 0;
        _lastOverlayWidthPx = 0;
        _lastOverlayHeightPx = 0;
        _hasLastDialogWindowRect = false;
        _hasLastDialogVisualRect = false;

        if (IsVisible)
        {
            if (_overlayHwnd != IntPtr.Zero)
            {
                NativeMethods.SetWindowLongPtrEx(_overlayHwnd, NativeMethods.GWL_HWNDPARENT, IntPtr.Zero);
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

        _positionRecomputeRequested = true;
        PositionToDialog();
    }

    public void Shutdown()
    {
        _allowClose = true;
        CancelAndDisposeRefreshCts();
        CancelAndDisposeSingleSelectCts();
        Close();
    }

    public void OnDialogContentChanged()
    {
        if (!IsVisible || _dialogHwnd == IntPtr.Zero || _dialogController is null)
        {
            return;
        }

        _ = ProbeDialogSnapshotAsync(forceProbe: true);
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

    private void RequestCandidatesRefresh(bool immediate)
    {
        lock (_refreshRequestSync)
        {
            if (immediate)
            {
                _refreshImmediateRequested = true;
            }
        }

        _ = Interlocked.Increment(ref _refreshRequestVersion);
        CancelAndDisposeRefreshCts();

        if (Interlocked.CompareExchange(ref _refreshLoopInFlight, 1, 0) != 0)
        {
            return;
        }

        _ = Dispatcher.BeginInvoke(
            DispatcherPriority.Input,
            new Action(() => _ = ProcessRefreshLoopAsync()));
    }

    private async Task ProcessRefreshLoopAsync()
    {
        try
        {
            while (true)
            {
                var latestVersion = Interlocked.Read(ref _refreshRequestVersion);
                if (latestVersion <= Interlocked.Read(ref _refreshProcessedVersion))
                {
                    return;
                }

                var immediate = ConsumeRefreshImmediateFlag();
                await RefreshCandidatesAsync(latestVersion, immediate);
                _ = Interlocked.Exchange(ref _refreshProcessedVersion, latestVersion);
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException(
                "OverlayWindow.ProcessRefreshLoopAsync",
                ex,
                throttle: TimeSpan.FromSeconds(2));
        }
        finally
        {
            _ = Interlocked.Exchange(ref _refreshLoopInFlight, 0);

            if (Interlocked.Read(ref _refreshRequestVersion) > Interlocked.Read(ref _refreshProcessedVersion) &&
                Interlocked.CompareExchange(ref _refreshLoopInFlight, 1, 0) == 0)
            {
                _ = Dispatcher.BeginInvoke(
                    DispatcherPriority.Input,
                    new Action(() => _ = ProcessRefreshLoopAsync()));
            }
        }
    }

    private bool ConsumeRefreshImmediateFlag()
    {
        lock (_refreshRequestSync)
        {
            var immediate = _refreshImmediateRequested;
            _refreshImmediateRequested = false;
            return immediate;
        }
    }

    private bool IsRefreshObsolete(long requestVersion, CancellationToken token)
    {
        return token.IsCancellationRequested ||
               requestVersion != Interlocked.Read(ref _refreshRequestVersion);
    }

    private async Task RefreshCandidatesAsync(long requestVersion, bool immediate = false)
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
            SetResultSummary(null);
            UpdateItems([]);
            return;
        }

        if (IsRefreshObsolete(requestVersion, token))
        {
            return;
        }

        var initialOutcome = await QueryCandidatesByModeAsync(mode, keyword, sort, InitialQueryResultLimit, token);
        if (initialOutcome is null || IsRefreshObsolete(requestVersion, token))
        {
            return;
        }

        var (initialFilteredItems, initialStatus) = ApplyTypeFilter(initialOutcome.Value.Items, initialOutcome.Value.StatusMessage, initialOutcome.Value.CanApplyTypeFilter);
        var (initialDisplayItems, initialTrimmed) = CapDisplayItems(initialFilteredItems);
        SetStatus(BuildStatusForMode(mode, initialStatus, initialTrimmed, expandedScan: false));
        UpdateSearchResultSummary(mode, initialDisplayItems.Count);
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

        if (IsRefreshObsolete(requestVersion, token))
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
                    allowedExtensions,
                    token);

                if (pagedOutcome is not null && !IsRefreshObsolete(requestVersion, token))
                {
                    SetStatus(BuildStatusForMode(mode, pagedOutcome.Value.StatusMessage, trimmed: false, expandedScan: true));
                    UpdateSearchResultSummary(mode, pagedOutcome.Value.Items.Count);
                    UpdateItems(pagedOutcome.Value.Items);
                }

                return;
            }
        }

        var expandedOutcome = await QueryCandidatesByModeAsync(mode, keyword, sort, expandedLimit, token);
        if (expandedOutcome is null || IsRefreshObsolete(requestVersion, token))
        {
            return;
        }

        var (expandedFilteredItems, expandedStatus) = ApplyTypeFilter(expandedOutcome.Value.Items, expandedOutcome.Value.StatusMessage, expandedOutcome.Value.CanApplyTypeFilter);
        var (expandedDisplayItems, expandedTrimmed) = CapDisplayItems(expandedFilteredItems);
        SetStatus(BuildStatusForMode(mode, expandedStatus, expandedTrimmed, expandedScan: true));
        UpdateSearchResultSummary(mode, expandedDisplayItems.Count);
        UpdateItems(expandedDisplayItems);
    }

    private async Task<(IReadOnlyList<FileCandidate> Items, string? StatusMessage)?> QueryTypeFilteredCandidatesByPagingAsync(
        OverlayMode mode,
        string keyword,
        EverythingSortOption sort,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken token)
    {
        var result = new List<FileCandidate>(DisplayResultLimit);
        var offset = 0;
        var exhausted = false;

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

            foreach (var candidate in pageResult.Items)
            {
                result.Add(candidate);
                if (result.Count >= DisplayResultLimit)
                {
                    return (result, null);
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

        var status = result.Count == 0 && exhausted
            ? "没有符合当前文件类型的结果。"
            : null;
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
                            : null;
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
                            : null;
                    }

                    break;
                case OverlayMode.Tray:
                    items = _trayRepository.Query(string.Empty);
                    if (_isFolderDialog)
                    {
                        items = items.Where(candidate => candidate.IsDirectory).ToList();
                    }

                    caption = _isFolderDialog ? "托盘文件夹" : "托盘文件";
                    statusMessage = null;
                    break;
                case OverlayMode.Explorer:
                    items = _explorerWindowService.GetOpenLocationsCached(string.Empty);
                    _explorerWindowService.EnsureSnapshotFreshAsync(force: false);
                    caption = "资源管理器路径";
                    statusMessage = items.Count == 0
                        ? "未检测到可用资源管理器路径。"
                        : null;
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
        bool canApplyDialogTypeFilter)
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

                if (items.Count == 0 && string.IsNullOrWhiteSpace(statusMessage))
                {
                    statusMessage = "没有符合当前文件类型的结果。";
                }
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
        var previousCount = _items.Count;
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
        var shouldRecalculateHeight =
            previousVisibility != ResultHost.Visibility ||
            previousCount != _items.Count ||
            _lastMeasuredStatusVisibility != StatusText.Visibility ||
            _lastMeasuredSummaryVisibility != ResultSummaryText.Visibility;

        if (shouldRecalculateHeight)
        {
            RecalculateWindowHeight();
        }

        if (changed ||
            previousVisibility != ResultHost.Visibility ||
            (shouldRecalculateHeight && Math.Abs(previousHeight - Height) > 0.5))
        {
            RequestPositionUpdate();
        }
    }

    private void RecalculateWindowHeight()
    {
        SizeToContent = SizeToContent.Height;
        UpdateLayout();

        var desiredHeight = ActualHeight;
        SizeToContent = SizeToContent.Manual;
        Height = desiredHeight > ExpandedHeightDip
            ? ExpandedHeightDip
            : desiredHeight;

        _lastMeasuredStatusVisibility = StatusText.Visibility;
        _lastMeasuredSummaryVisibility = ResultSummaryText.Visibility;
    }

    private void PositionToDialog()
    {
        if (_dialogHwnd == IntPtr.Zero)
        {
            return;
        }

        if (!NativeMethods.IsWindow(_dialogHwnd) || !NativeMethods.IsWindowVisible(_dialogHwnd))
        {
            DetachDialog();
            return;
        }

        if (!NativeMethods.GetWindowRect(_dialogHwnd, out var rect))
        {
            return;
        }

        PositionToDialog(rect);
    }

    private void PositionToDialog(NativeMethods.RECT windowRect)
    {
        var windowRectChanged = !_hasLastDialogWindowRect || !RectEquals(_lastDialogWindowRect, windowRect);
        _lastDialogWindowRect = windowRect;
        _hasLastDialogWindowRect = true;

        NativeMethods.RECT rect;
        if (windowRectChanged || !_hasLastDialogVisualRect)
        {
            rect = windowRect;

            // Win10/11 file dialogs include an invisible resize border in GetWindowRect.
            // Use the visual frame bounds when available to remove the alignment gap.
            if (NativeMethods.DwmGetWindowAttribute(
                    _dialogHwnd,
                    NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS,
                    out var extRect,
                    Marshal.SizeOf(typeof(NativeMethods.RECT))) == 0 &&
                extRect.Width > 0 &&
                extRect.Height > 0)
            {
                rect = extRect;
            }

            _lastDialogVisualRect = rect;
            _hasLastDialogVisualRect = true;
        }
        else
        {
            rect = _lastDialogVisualRect;
        }

        _positionRecomputeRequested = false;

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

        if (_hasAppliedPosition &&
            _lastOverlayLeftPx == desiredLeftPx &&
            _lastOverlayTopPx == desiredTopPx &&
            _lastOverlayWidthPx == desiredWidthPx &&
            _lastOverlayHeightPx == finalHeightPx)
        {
            return;
        }

        _hasAppliedPosition = true;
        _lastOverlayLeftPx = desiredLeftPx;
        _lastOverlayTopPx = desiredTopPx;
        _lastOverlayWidthPx = desiredWidthPx;
        _lastOverlayHeightPx = finalHeightPx;

        var overlayHwnd = _overlayHwnd != IntPtr.Zero
            ? _overlayHwnd
            : new WindowInteropHelper(this).Handle;
        if (overlayHwnd != IntPtr.Zero)
        {
            _ = NativeMethods.SetWindowPos(
                overlayHwnd,
                IntPtr.Zero,
                desiredLeftPx,
                desiredTopPx,
                desiredWidthPx,
                finalHeightPx,
                NativeMethods.SWP_NOACTIVATE |
                NativeMethods.SWP_NOZORDER |
                NativeMethods.SWP_NOSENDCHANGING);
        }
        else
        {
            Width = PxToDip(desiredWidthPx, scale);
            Height = finalHeightDip;
            Left = PxToDip(desiredLeftPx, scale);
            Top = PxToDip(desiredTopPx, scale);
        }
    }

    private static bool RectEquals(NativeMethods.RECT left, NativeMethods.RECT right)
    {
        return left.Left == right.Left &&
               left.Top == right.Top &&
               left.Right == right.Right &&
               left.Bottom == right.Bottom;
    }

    private void UpdateResultHostHeight(double windowHeightDip)
    {
        // No longer forcing ResultHost to a specific height.
        // XAML Grid RowDefinitions takes care of it natively.
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

    private void SetResultSummary(string? message)
    {
        var previousVisibility = ResultSummaryText.Visibility;
        var previousText = ResultSummaryText.Text;

        if (string.IsNullOrWhiteSpace(message))
        {
            ResultSummaryText.Visibility = Visibility.Collapsed;
            ResultSummaryText.Text = string.Empty;
        }
        else
        {
            ResultSummaryText.Visibility = Visibility.Visible;
            ResultSummaryText.Text = message;
        }

        if ((previousVisibility != ResultSummaryText.Visibility ||
             !string.Equals(previousText, ResultSummaryText.Text, StringComparison.Ordinal)) &&
            ResultHost.Visibility == Visibility.Visible)
        {
            RequestPositionUpdate();
        }
    }

    private void UpdateSearchResultSummary(OverlayMode mode, int displayCount)
    {
        if (mode != OverlayMode.Search || displayCount <= 0)
        {
            SetResultSummary(null);
            return;
        }

        SetResultSummary($"搜索查询结果 {displayCount} 条");
    }

    private static string? BuildStatusForMode(OverlayMode mode, string? statusMessage, bool trimmed, bool expandedScan)
    {
        if (string.IsNullOrWhiteSpace(statusMessage))
        {
            return null;
        }

        if (mode == OverlayMode.Search)
        {
            return statusMessage;
        }

        return BuildStatus(statusMessage, trimmed, expandedScan);
    }

    private EverythingSortOption CurrentSort =>
        _sortOptions[Math.Clamp(_sortIndex, 0, _sortOptions.Count - 1)].Value;

    private bool IsTypeFilterActive => _preferredTypeFilterEnabled && !_isFolderDialog;

    private void SearchDebounceTimer_OnTick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        RequestCandidatesRefresh(immediate: false);
    }

    private async void DialogMonitorTimer_OnTick(object? sender, EventArgs e)
    {
        await ProbeDialogSnapshotAsync(forceProbe: false);
    }

    private async Task ProbeDialogSnapshotAsync(bool forceProbe)
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
            var shouldPollFilterText = ShouldPollDialogFilterText();
            var shouldProbeFolderState = forceProbe || ShouldProbeFolderState();
            if (!forceProbe && !ShouldProbeDialogSnapshot(shouldPollFilterText, shouldProbeFolderState))
            {
                return;
            }

            (bool IsFolderDialog, string FilterText) snapshot;
            try
            {
                snapshot = await Task.Run(
                        () => CaptureDialogSnapshot(
                            controller,
                            shouldPollFilterText,
                            _lastFileTypeFilterText,
                            shouldProbeFolderState,
                            _isFolderDialog))
                    .WaitAsync(DialogMonitorTimeout);
            }
            catch (TimeoutException)
            {
                return;
            }

            if (_dialogController != controller)
            {
                return;
            }

            _lastDialogSnapshotProbeUtc = DateTime.UtcNow;
            if (shouldProbeFolderState)
            {
                _lastFolderStateProbeUtc = _lastDialogSnapshotProbeUtc;
            }

            var folderModeChanged = snapshot.IsFolderDialog != _isFolderDialog;
            if (folderModeChanged)
            {
                _isFolderDialog = snapshot.IsFolderDialog;
                ApplyModeUiState();
            }

            var filterChanged = false;
            if (shouldPollFilterText)
            {
                filterChanged = !string.Equals(snapshot.FilterText, _lastFileTypeFilterText, StringComparison.OrdinalIgnoreCase);
                if (filterChanged)
                {
                    _lastFileTypeFilterText = snapshot.FilterText;
                }
            }

            if (folderModeChanged || filterChanged)
            {
                _searchDebounceTimer.Stop();
                RequestCandidatesRefresh(immediate: true);
            }
        }
        catch (Exception ex)
        {
            AppLogger.LogException(
                "OverlayWindow.ProbeDialogSnapshotAsync",
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
        if (_currentMode == OverlayMode.Explorer &&
            !string.IsNullOrWhiteSpace(SearchBox.Text))
        {
            _currentMode = OverlayMode.Search;
            ApplyModeUiState();
        }

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
        SyncSortMenuSelectionState();
        _sortContextMenu.PlacementTarget = SortButton;
        _sortContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        _sortContextMenu.IsOpen = true;
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
        RequestCandidatesRefresh(immediate: true);
    }

    private void ResetModeForDialog()
    {
        _explorerWindowService.EnsureSnapshotFreshAsync(force: false);
        _currentMode = _explorerWindowService.HasOpenLocationsCached()
            ? OverlayMode.Explorer
            : OverlayMode.Search;
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
        UpdateDialogMonitorInterval();
    }

    private void UpdateSortButtonLabel()
    {
        // SortLabelText.Text = _sortOptions[_sortIndex].Label;
        SortButton.ToolTip = $"点击切换搜索排序 ({_sortOptions[_sortIndex].Label})";
    }

    private ContextMenu CreateSortContextMenu()
    {
        var contextMenu = new ContextMenu();
        for (var i = 0; i < _sortOptions.Count; i++)
        {
            var option = _sortOptions[i];
            var item = new MenuItem
            {
                Header = option.Label,
                Tag = i,
            };
            item.Click += SortMenuItem_OnClick;
            contextMenu.Items.Add(item);
        }

        return contextMenu;
    }

    private void SortMenuItem_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: int index } || _sortIndex == index)
        {
            return;
        }

        _sortIndex = index;
        UpdateSortButtonLabel();

        if (_currentMode is OverlayMode.Search or OverlayMode.Recent)
        {
            _searchDebounceTimer.Stop();
            RequestCandidatesRefresh(immediate: true);
        }
    }

    private void SyncSortMenuSelectionState()
    {
        for (var i = 0; i < _sortContextMenu.Items.Count; i++)
        {
            if (_sortContextMenu.Items[i] is MenuItem item)
            {
                item.IsChecked = i == _sortIndex;
            }
        }
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
        RequestCandidatesRefresh(immediate: true);
    }

    private void ConfigureDropInteropCompatibility(IntPtr overlayHwnd)
    {
        if (_dropInteropConfigured || overlayHwnd == IntPtr.Zero)
        {
            return;
        }

        TryAllowDropMessage(overlayHwnd, NativeMethods.WM_DROPFILES);
        TryAllowDropMessage(overlayHwnd, NativeMethods.WM_COPYDATA);
        TryAllowDropMessage(overlayHwnd, NativeMethods.WM_COPYGLOBALDATA);
        NativeMethods.DragAcceptFiles(overlayHwnd, true);
        _dropInteropConfigured = true;
    }

    private static void TryAllowDropMessage(IntPtr hwnd, uint message)
    {
        if (NativeMethods.ChangeWindowMessageFilterEx(hwnd, message, NativeMethods.MSGFLT_ALLOW, IntPtr.Zero))
        {
            return;
        }

        var error = Marshal.GetLastWin32Error();
        if (error != 0)
        {
            AppLogger.LogWarning($"ChangeWindowMessageFilterEx failed. message=0x{message:X}; win32={error}");
        }
    }

    private void Window_OnPreviewDrop(object sender, System.Windows.DragEventArgs e)
    {
        var paths = DropPathExtractor.ExtractDroppedPaths(e.Data);
        if (paths.Count == 0)
        {
            SetStatus("未识别到可加入托盘的文件。");
            return;
        }

        var added = _trayRepository.AddMany(paths);
        SetStatus(added > 0 ? $"已加入托盘：{added} 项。" : "拖入项已在托盘中。");

        if (_currentMode == OverlayMode.Tray)
        {
            RequestCandidatesRefresh(immediate: true);
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
        if (sender is ListBoxItem { DataContext: FileCandidate candidate } item)
        {
            if (_currentMode == OverlayMode.Tray && item.ContextMenu is { } menu)
            {
                menu.PlacementTarget = item;
                menu.IsOpen = true;
            }
            else
            {
                if (!ReferenceEquals(ResultList.SelectedItem, candidate))
                {
                    ResultList.SelectedItem = candidate;
                }
                _ = ExecuteSingleAsync(candidate);
            }
            e.Handled = true;
        }
    }

    private void ListBoxItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBoxItem { DataContext: FileCandidate candidate })
        {
            return;
        }

        if (!ReferenceEquals(ResultList.SelectedItem, candidate))
        {
            ResultList.SelectedItem = candidate;
        }

        _ = ExecuteConfirmAsync(candidate);
        e.Handled = true;
    }

    private void TrayMenuItem_Pin_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem && 
            menuItem.DataContext is FileCandidate candidate)
        {
            _trayRepository.PinToTop(candidate.FullPath);
            RequestCandidatesRefresh(immediate: true);
        }
    }

    private void TrayMenuItem_Delete_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem && 
            menuItem.DataContext is FileCandidate candidate)
        {
            _trayRepository.Remove(candidate.FullPath);
            RequestCandidatesRefresh(immediate: true);
        }
    }

    private void Window_OnPreviewDragOver(object sender, System.Windows.DragEventArgs e)
    {
        e.Effects = DropPathExtractor.HasDroppedFiles(e.Data)
            ? System.Windows.DragDropEffects.Copy
            : System.Windows.DragDropEffects.None;
        e.Handled = true;
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

        return allowedExtensions.Contains(extension);
    }

    private bool ShouldPollDialogFilterText()
    {
        return IsTypeFilterActive && _currentMode is OverlayMode.Search or OverlayMode.Recent;
    }

    private bool ShouldProbeFolderState()
    {
        if (_lastFolderStateProbeUtc == DateTime.MinValue)
        {
            return true;
        }

        return DateTime.UtcNow - _lastFolderStateProbeUtc >= FolderStateProbeInterval;
    }

    private bool ShouldProbeDialogSnapshot(bool shouldPollFilterText, bool shouldProbeFolderState)
    {
        if (shouldProbeFolderState)
        {
            return true;
        }

        var interval = shouldPollFilterText
            ? DialogMonitorBusyInterval
            : DialogMonitorIdleInterval;

        if (_lastDialogSnapshotProbeUtc == DateTime.MinValue)
        {
            return true;
        }

        return DateTime.UtcNow - _lastDialogSnapshotProbeUtc >= interval;
    }

    private void UpdateDialogMonitorInterval()
    {
        _dialogMonitorTimer.Interval = DialogMonitorLivenessInterval;
    }

    private static (bool IsFolderDialog, string FilterText) CaptureDialogSnapshot(
        FileDialogAutomationController controller,
        bool includeFilterText,
        string fallbackFilterText,
        bool includeFolderState,
        bool fallbackFolderState)
    {
        var isFolderDialog = includeFolderState
            ? controller.IsFolderSelectionDialog()
            : fallbackFolderState;
        if (!includeFilterText)
        {
            return (isFolderDialog, fallbackFilterText);
        }

        return (isFolderDialog, controller.GetCurrentFileTypeFilterText());
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
