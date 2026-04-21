using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using SuperSelect.App.Models;

namespace SuperSelect.App.Services;

internal sealed class EverythingService
{
    internal readonly record struct BlockedPathEntry(
        string FullPath,
        bool IsDirectory,
        bool BlockSearch,
        bool BlockRecent);

    private sealed class BlocklistPayload
    {
        public List<BlockRulePayload> Rules { get; set; } = [];
    }

    private sealed class BlockRulePayload
    {
        public string Path { get; set; } = string.Empty;
        public bool IsDirectory { get; set; }
        public bool BlockSearch { get; set; }
        public bool BlockRecent { get; set; }
    }

    private sealed class BlockRuleState
    {
        public bool IsDirectory { get; set; }
        public bool BlockSearch { get; set; }
        public bool BlockRecent { get; set; }
    }

    private readonly object _syncRoot = new();
    private readonly string _blocklistPath;
    private static readonly TimeSpan AvailabilityProbeInterval = TimeSpan.FromMilliseconds(500);
    private bool _isAvailable = true;
    private string _lastErrorMessage = string.Empty;
    private DateTime _lastAvailabilityProbeUtc = DateTime.MinValue;
    private Dictionary<string, BlockRuleState> _blockedPathRules = new(StringComparer.OrdinalIgnoreCase);
    private string _searchBlockQuerySuffix = string.Empty;
    private string _recentBlockQuerySuffix = string.Empty;
    [ThreadStatic]
    private static StringBuilder? _threadPathBuilder;

    public EverythingService()
    {
        var storeDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SuperSelect");
        Directory.CreateDirectory(storeDirectory);
        _blocklistPath = Path.Combine(storeDirectory, "recent-blocklist.json");
        LoadBlocklist();
    }

    public bool IsAvailable
    {
        get
        {
            lock (_syncRoot)
            {
                return _isAvailable;
            }
        }
    }

    public string LastErrorMessage
    {
        get
        {
            lock (_syncRoot)
            {
                return _lastErrorMessage;
            }
        }
    }

    public Task<IReadOnlyList<FileCandidate>> SearchFilesAsync(
        string keyword,
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = BuildSearchQuery(isDirectoryQuery: false, keyword);

        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: false, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> SearchFoldersAsync(
        string keyword,
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = BuildSearchQuery(isDirectoryQuery: true, keyword);

        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: true, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> SearchEntriesAsync(
        string keyword,
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = BuildSearchMixedQuery(keyword);

        return Task.Run(
            () => QueryInternalMixed(query, ToEverythingSort(sortOption), maxResults, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFilesAsync(
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = BuildRecentQuery("file:");
        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: false, CandidateSource.EverythingRecent, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFoldersAsync(
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = BuildRecentQuery("folder:");
        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: true, CandidateSource.EverythingRecent, cancellationToken),
            cancellationToken);
    }

    public Task<(IReadOnlyList<FileCandidate> Items, int RawCount)> SearchFilesFilteredPageAsync(
        string keyword,
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken cancellationToken)
    {
        var query = BuildSearchQuery(isDirectoryQuery: false, keyword);

        return Task.Run(
            () => QueryPageFilteredInternal(
                query,
                ToEverythingSort(sortOption),
                offset,
                pageSize,
                CandidateSource.EverythingSearch,
                allowedExtensions,
                cancellationToken),
            cancellationToken);
    }

    public Task<(IReadOnlyList<FileCandidate> Items, int RawCount)> RecentFilesFilteredPageAsync(
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken cancellationToken)
    {
        var query = BuildRecentQuery("file:");
        return Task.Run(
            () => QueryPageFilteredInternal(
                query,
                ToEverythingSort(sortOption),
                offset,
                pageSize,
                CandidateSource.EverythingRecent,
                allowedExtensions,
                cancellationToken),
            cancellationToken);
    }

    public bool BlockSearchFile(string fullPath)
    {
        return UpdatePathBlockState(fullPath, isDirectory: false, blockSearch: true, blockRecent: null);
    }

    public bool BlockSearchDirectory(string directoryPath)
    {
        return UpdatePathBlockState(directoryPath, isDirectory: true, blockSearch: true, blockRecent: null);
    }

    public bool BlockRecentFile(string fullPath)
    {
        return UpdatePathBlockState(fullPath, isDirectory: false, blockSearch: null, blockRecent: true);
    }

    public bool BlockRecentDirectory(string directoryPath)
    {
        return UpdatePathBlockState(directoryPath, isDirectory: true, blockSearch: null, blockRecent: true);
    }

    public bool UnblockSearchFile(string fullPath)
    {
        return UpdatePathBlockState(fullPath, isDirectory: false, blockSearch: false, blockRecent: null);
    }

    public bool UnblockSearchDirectory(string directoryPath)
    {
        return UpdatePathBlockState(directoryPath, isDirectory: true, blockSearch: false, blockRecent: null);
    }

    public bool UnblockRecentFile(string fullPath)
    {
        return UpdatePathBlockState(fullPath, isDirectory: false, blockSearch: null, blockRecent: false);
    }

    public bool UnblockRecentDirectory(string directoryPath)
    {
        return UpdatePathBlockState(directoryPath, isDirectory: true, blockSearch: null, blockRecent: false);
    }

    public bool SetBlockedScopes(string fullPath, bool isDirectory, bool blockSearch, bool blockRecent)
    {
        return UpdatePathBlockState(fullPath, isDirectory, blockSearch, blockRecent);
    }

    public IReadOnlyList<BlockedPathEntry> GetBlockedPathEntries()
    {
        lock (_syncRoot)
        {
            return _blockedPathRules
                .OrderBy(static pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                .Select(static pair => new BlockedPathEntry(
                    pair.Key,
                    pair.Value.IsDirectory,
                    pair.Value.BlockSearch,
                    pair.Value.BlockRecent))
                .ToList();
        }
    }

    private bool UpdatePathBlockState(string fullPath, bool isDirectory, bool? blockSearch, bool? blockRecent)
    {
        var normalized = NormalizePath(fullPath);
        if (normalized is null)
        {
            return false;
        }

        lock (_syncRoot)
        {
            return UpdatePathBlockStateLocked(normalized, isDirectory, blockSearch, blockRecent);
        }
    }

    private bool UpdatePathBlockStateLocked(string normalizedPath, bool isDirectory, bool? blockSearch, bool? blockRecent)
    {
        _ = _blockedPathRules.TryGetValue(normalizedPath, out var existingState);
        existingState ??= new BlockRuleState
        {
            IsDirectory = isDirectory,
            BlockSearch = false,
            BlockRecent = false,
        };

        var nextIsDirectory = existingState.IsDirectory || isDirectory;
        var nextBlockSearch = blockSearch ?? existingState.BlockSearch;
        var nextBlockRecent = blockRecent ?? existingState.BlockRecent;

        if (!nextBlockSearch && !nextBlockRecent)
        {
            if (!_blockedPathRules.Remove(normalizedPath))
            {
                return false;
            }

            RebuildBlockQuerySuffixLocked();
            SaveBlocklistLocked();
            return true;
        }

        if (_blockedPathRules.TryGetValue(normalizedPath, out var currentState) &&
            currentState.IsDirectory == nextIsDirectory &&
            currentState.BlockSearch == nextBlockSearch &&
            currentState.BlockRecent == nextBlockRecent)
        {
            return false;
        }

        _blockedPathRules[normalizedPath] = new BlockRuleState
        {
            IsDirectory = nextIsDirectory,
            BlockSearch = nextBlockSearch,
            BlockRecent = nextBlockRecent,
        };

        RebuildBlockQuerySuffixLocked();
        SaveBlocklistLocked();
        return true;
    }

    private string BuildSearchMixedQuery(string keyword)
    {
        var baseQuery = string.IsNullOrWhiteSpace(keyword)
            ? string.Empty
            : keyword.Trim();
        return AppendBlockSuffix(baseQuery, forSearchScope: true);
    }

    private string BuildSearchQuery(bool isDirectoryQuery, string keyword)
    {
        var prefix = isDirectoryQuery ? "folder:" : "file:";
        var baseQuery = string.IsNullOrWhiteSpace(keyword)
            ? prefix
            : $"{prefix} {keyword.Trim()}";
        return AppendBlockSuffix(baseQuery, forSearchScope: true);
    }

    private string BuildRecentQuery(string baseQuery)
    {
        return AppendBlockSuffix(baseQuery, forSearchScope: false);
    }

    private string AppendBlockSuffix(string baseQuery, bool forSearchScope)
    {
        var suffix = forSearchScope ? _searchBlockQuerySuffix : _recentBlockQuerySuffix;
        if (string.IsNullOrEmpty(suffix))
        {
            return baseQuery;
        }

        if (string.IsNullOrWhiteSpace(baseQuery))
        {
            return suffix[0] == ' ' ? suffix[1..] : suffix;
        }

        return baseQuery + suffix;
    }

    private static void AppendPathExclusion(StringBuilder builder, string path, bool isDirectory)
    {
        builder.Append(' ');
        builder.Append("!path:\"");
        builder.Append(isDirectory ? EnsureTrailingDirectorySeparator(path) : path);
        builder.Append('"');
    }

    private void RebuildBlockQuerySuffixLocked()
    {
        var searchBuilder = new StringBuilder();
        var recentBuilder = new StringBuilder();

        foreach (var pair in _blockedPathRules)
        {
            var state = pair.Value;
            if (state.BlockSearch)
            {
                AppendPathExclusion(searchBuilder, pair.Key, state.IsDirectory);
            }

            if (state.BlockRecent)
            {
                AppendPathExclusion(recentBuilder, pair.Key, state.IsDirectory);
            }
        }

        _searchBlockQuerySuffix = searchBuilder.ToString();
        _recentBlockQuerySuffix = recentBuilder.ToString();
    }

    private IReadOnlyList<FileCandidate> QueryInternal(
        string query,
        uint sortType,
        int maxResults,
        bool isDirectoryResult,
        CandidateSource source,
        CancellationToken cancellationToken)
    {
        return QueryPageInternal(
            query,
            sortType,
            offset: 0,
            maxResults: maxResults,
            isDirectoryResult: isDirectoryResult,
            source: source,
            cancellationToken: cancellationToken);
    }

    private IReadOnlyList<FileCandidate> QueryInternalMixed(
        string query,
        uint sortType,
        int maxResults,
        CandidateSource source,
        CancellationToken cancellationToken)
    {
        return QueryPageInternalMixed(
            query,
            sortType,
            offset: 0,
            maxResults: maxResults,
            source: source,
            cancellationToken: cancellationToken);
    }

    private IReadOnlyList<FileCandidate> QueryPageInternal(
        string query,
        uint sortType,
        int offset,
        int maxResults,
        bool isDirectoryResult,
        CandidateSource source,
        CancellationToken cancellationToken)
    {
        var boundedOffset = Math.Max(0, offset);
        var boundedMaxResults = Math.Clamp(maxResults, 1, 20000);

        lock (_syncRoot)
        {
            if (!EnsureAvailableLocked())
            {
                return [];
            }

            try
            {
                ConfigureEverythingQuery(query, sortType, boundedOffset, boundedMaxResults);

                if (!Native.Everything_QueryW(true))
                {
                    var errorCode = Native.Everything_GetLastError();
                    _lastErrorMessage = $"查询失败（错误码: {errorCode}）";
                    return [];
                }

                var count = (int)Native.Everything_GetNumResults();
                var result = new List<FileCandidate>(Math.Min(count, boundedMaxResults));
                var builder = GetPathBuilder();

                for (var i = 0; i < count; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        return [];
                    }

                    builder.Clear();
                    Native.Everything_GetResultFullPathNameW((uint)i, builder, (uint)builder.Capacity);

                    var fullPath = builder.ToString();
                    if (string.IsNullOrWhiteSpace(fullPath))
                    {
                        continue;
                    }

                    result.Add(CreateCandidate(fullPath, isDirectoryResult, source));
                }

                return result;
            }
            catch (DllNotFoundException)
            {
                SetUnavailableLocked("Everything64.dll 未找到。请把 SDK DLL 放到可执行程序目录，或加入系统 PATH。");
                return [];
            }
            catch (EntryPointNotFoundException)
            {
                SetUnavailableLocked("Everything SDK 接口不可用。请确认 DLL 版本与程序架构一致（x64）。");
                return [];
            }
            catch (Exception ex)
            {
                _lastErrorMessage = ex.Message;
                return [];
            }
        }
    }

    private IReadOnlyList<FileCandidate> QueryPageInternalMixed(
        string query,
        uint sortType,
        int offset,
        int maxResults,
        CandidateSource source,
        CancellationToken cancellationToken)
    {
        var boundedOffset = Math.Max(0, offset);
        var boundedMaxResults = Math.Clamp(maxResults, 1, 20000);

        lock (_syncRoot)
        {
            if (!EnsureAvailableLocked())
            {
                return [];
            }

            try
            {
                ConfigureEverythingQuery(query, sortType, boundedOffset, boundedMaxResults);

                if (!Native.Everything_QueryW(true))
                {
                    var errorCode = Native.Everything_GetLastError();
                    _lastErrorMessage = $"查询失败（错误码: {errorCode}）";
                    return [];
                }

                var count = (int)Native.Everything_GetNumResults();
                var result = new List<FileCandidate>(Math.Min(count, boundedMaxResults));
                var builder = GetPathBuilder();

                for (var i = 0; i < count; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        return [];
                    }

                    builder.Clear();
                    Native.Everything_GetResultFullPathNameW((uint)i, builder, (uint)builder.Capacity);

                    var fullPath = builder.ToString();
                    if (string.IsNullOrWhiteSpace(fullPath))
                    {
                        continue;
                    }

                    var isDirectory = Native.Everything_IsFolderResult((uint)i);
                    result.Add(CreateCandidate(fullPath, isDirectory, source));
                }

                return result;
            }
            catch (DllNotFoundException)
            {
                SetUnavailableLocked("Everything64.dll 未找到。请把 SDK DLL 放到可执行程序目录，或加入系统 PATH。");
                return [];
            }
            catch (EntryPointNotFoundException)
            {
                SetUnavailableLocked("Everything SDK 接口不可用。请确认 DLL 版本与程序架构一致（x64）。");
                return [];
            }
            catch (Exception ex)
            {
                _lastErrorMessage = ex.Message;
                return [];
            }
        }
    }

    private (IReadOnlyList<FileCandidate> Items, int RawCount) QueryPageFilteredInternal(
        string query,
        uint sortType,
        int offset,
        int maxResults,
        CandidateSource source,
        IReadOnlySet<string> allowedExtensions,
        CancellationToken cancellationToken)
    {
        var boundedOffset = Math.Max(0, offset);
        var boundedMaxResults = Math.Clamp(maxResults, 1, 20000);
        var canFilterByExtensions = allowedExtensions.Count > 0;

        lock (_syncRoot)
        {
            if (!EnsureAvailableLocked())
            {
                return ([], 0);
            }

            try
            {
                ConfigureEverythingQuery(query, sortType, boundedOffset, boundedMaxResults);

                if (!Native.Everything_QueryW(true))
                {
                    var errorCode = Native.Everything_GetLastError();
                    _lastErrorMessage = $"查询失败（错误码: {errorCode}）";
                    return ([], 0);
                }

                var rawCount = (int)Native.Everything_GetNumResults();
                var result = new List<FileCandidate>(Math.Min(rawCount, boundedMaxResults));
                var builder = GetPathBuilder();

                for (var i = 0; i < rawCount; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        return ([], 0);
                    }

                    builder.Clear();
                    Native.Everything_GetResultFullPathNameW((uint)i, builder, (uint)builder.Capacity);

                    var fullPath = builder.ToString();
                    if (string.IsNullOrWhiteSpace(fullPath))
                    {
                        continue;
                    }

                    if (canFilterByExtensions && !HasAllowedExtension(fullPath, allowedExtensions))
                    {
                        continue;
                    }

                    result.Add(CreateCandidate(fullPath, isDirectory: false, source));
                }

                return (result, rawCount);
            }
            catch (DllNotFoundException)
            {
                SetUnavailableLocked("Everything64.dll 未找到。请把 SDK DLL 放到可执行程序目录，或加入系统 PATH。");
                return ([], 0);
            }
            catch (EntryPointNotFoundException)
            {
                SetUnavailableLocked("Everything SDK 接口不可用。请确认 DLL 版本与程序架构一致（x64）。");
                return ([], 0);
            }
            catch (Exception ex)
            {
                _lastErrorMessage = ex.Message;
                return ([], 0);
            }
        }
    }

    private static void ConfigureEverythingQuery(string query, uint sortType, int offset, int maxResults)
    {
        Native.Everything_SetMatchPath(false);
        Native.Everything_SetMatchCase(false);
        Native.Everything_SetRegex(false);
        Native.Everything_SetRequestFlags(RequestFlags.FullPathAndFileName);
        Native.Everything_SetSort(sortType);
        Native.Everything_SetOffset((uint)offset);
        Native.Everything_SetMax((uint)maxResults);
        Native.Everything_SetSearchW(query);
    }

    private static FileCandidate CreateCandidate(string fullPath, bool isDirectory, CandidateSource source)
    {
        var displayName = Path.GetFileName(fullPath);
        if (string.IsNullOrWhiteSpace(displayName))
        {
            displayName = fullPath;
        }

        return new FileCandidate
        {
            FullPath = fullPath,
            DisplayName = displayName,
            SecondaryText = Path.GetDirectoryName(fullPath) ?? string.Empty,
            IsDirectory = isDirectory,
            Source = source,
        };
    }

    private static StringBuilder GetPathBuilder()
    {
        var builder = _threadPathBuilder;
        if (builder is null)
        {
            builder = new StringBuilder(4096);
            _threadPathBuilder = builder;
        }
        else if (builder.Capacity < 4096)
        {
            _ = builder.EnsureCapacity(4096);
        }

        return builder;
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

    private static string EnsureTrailingDirectorySeparator(string directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            return directoryPath;
        }

        return directoryPath.EndsWith(Path.DirectorySeparatorChar) ||
               directoryPath.EndsWith(Path.AltDirectorySeparatorChar)
            ? directoryPath
            : directoryPath + Path.DirectorySeparatorChar;
    }

    private static string? NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            var normalized = Path.GetFullPath(path.Trim().Trim('"'));
            return normalized;
        }
        catch
        {
            return null;
        }
    }

    private void LoadBlocklist()
    {
        try
        {
            if (!File.Exists(_blocklistPath))
            {
                return;
            }

            var payload = JsonSerializer.Deserialize<BlocklistPayload>(File.ReadAllText(_blocklistPath));
            if (payload is null)
            {
                return;
            }

            var loadedRules = new Dictionary<string, BlockRuleState>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < payload.Rules.Count; i++)
            {
                var rule = payload.Rules[i];
                if (!rule.BlockSearch && !rule.BlockRecent)
                {
                    continue;
                }

                var normalized = NormalizePath(rule.Path);
                if (normalized is null)
                {
                    continue;
                }

                if (!loadedRules.TryGetValue(normalized, out var state))
                {
                    loadedRules[normalized] = new BlockRuleState
                    {
                        IsDirectory = rule.IsDirectory,
                        BlockSearch = rule.BlockSearch,
                        BlockRecent = rule.BlockRecent,
                    };
                    continue;
                }

                state.IsDirectory = state.IsDirectory || rule.IsDirectory;
                state.BlockSearch = state.BlockSearch || rule.BlockSearch;
                state.BlockRecent = state.BlockRecent || rule.BlockRecent;
            }

            lock (_syncRoot)
            {
                _blockedPathRules = loadedRules;
                RebuildBlockQuerySuffixLocked();
            }
        }
        catch
        {
            // Ignore malformed blocklist and continue with empty rules.
        }
    }

    private void SaveBlocklistLocked()
    {
        try
        {
            var payload = new BlocklistPayload
            {
                Rules = _blockedPathRules
                    .OrderBy(static pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(static pair => new BlockRulePayload
                    {
                        Path = pair.Key,
                        IsDirectory = pair.Value.IsDirectory,
                        BlockSearch = pair.Value.BlockSearch,
                        BlockRecent = pair.Value.BlockRecent,
                    })
                    .ToList(),
            };

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_blocklistPath, json);
        }
        catch
        {
            // Ignore persistence failure.
        }
    }

    private bool EnsureAvailableLocked()
    {
        if (!_isAvailable)
        {
            return false;
        }

        var nowUtc = DateTime.UtcNow;
        if (_lastAvailabilityProbeUtc != DateTime.MinValue &&
            nowUtc - _lastAvailabilityProbeUtc < AvailabilityProbeInterval)
        {
            return true;
        }

        _lastAvailabilityProbeUtc = nowUtc;

        try
        {
            _ = Native.Everything_GetMajorVersion();
            if (!Native.Everything_IsDBLoaded())
            {
                _lastErrorMessage = "Everything 数据库尚未加载完成。";
            }

            return true;
        }
        catch (DllNotFoundException)
        {
            SetUnavailableLocked("Everything64.dll 未找到。");
            return false;
        }
        catch (EntryPointNotFoundException)
        {
            SetUnavailableLocked("Everything SDK 函数入口不存在。");
            return false;
        }
    }

    private void SetUnavailableLocked(string message)
    {
        _isAvailable = false;
        _lastErrorMessage = message;
        _lastAvailabilityProbeUtc = DateTime.MinValue;
    }

    private static uint ToEverythingSort(EverythingSortOption sortOption)
    {
        return sortOption switch
        {
            EverythingSortOption.NameAsc => EverythingSort.NameAscending,
            EverythingSortOption.NameDesc => EverythingSort.NameDescending,
            EverythingSortOption.PathAsc => EverythingSort.PathAscending,
            EverythingSortOption.PathDesc => EverythingSort.PathDescending,
            EverythingSortOption.DateModifiedAsc => EverythingSort.DateModifiedAscending,
            _ => EverythingSort.DateModifiedDescending,
        };
    }

    private static class RequestFlags
    {
        internal const uint FullPathAndFileName = 0x00000004;
    }

    private static class EverythingSort
    {
        internal const uint NameAscending = 1;
        internal const uint NameDescending = 2;
        internal const uint PathAscending = 3;
        internal const uint PathDescending = 4;
        internal const uint DateModifiedAscending = 13;
        internal const uint DateModifiedDescending = 14;
    }

    private static class Native
    {
        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        internal static extern void Everything_SetSearchW(string searchString);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetMatchPath([MarshalAs(UnmanagedType.Bool)] bool enable);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetMatchCase([MarshalAs(UnmanagedType.Bool)] bool enable);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetRegex([MarshalAs(UnmanagedType.Bool)] bool enable);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetRequestFlags(uint requestFlags);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetSort(uint sortType);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetOffset(uint offset);

        [DllImport("Everything64.dll")]
        internal static extern void Everything_SetMax(uint max);

        [DllImport("Everything64.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool Everything_QueryW([MarshalAs(UnmanagedType.Bool)] bool wait);

        [DllImport("Everything64.dll")]
        internal static extern uint Everything_GetNumResults();

        [DllImport("Everything64.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool Everything_IsFolderResult(uint index);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        internal static extern void Everything_GetResultFullPathNameW(uint index, StringBuilder result, uint maxCount);

        [DllImport("Everything64.dll")]
        internal static extern uint Everything_GetLastError();

        [DllImport("Everything64.dll")]
        internal static extern uint Everything_GetMajorVersion();

        [DllImport("Everything64.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool Everything_IsDBLoaded();
    }
}
