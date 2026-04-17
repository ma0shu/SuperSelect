using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using SuperSelect.App.Models;

namespace SuperSelect.App.Services;

internal sealed class EverythingService
{
    private readonly object _syncRoot = new();
    private bool _isAvailable = true;
    private string _lastErrorMessage = string.Empty;
    [ThreadStatic]
    private static StringBuilder? _threadPathBuilder;

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
        var query = string.IsNullOrWhiteSpace(keyword)
            ? "file:"
            : $"file: {keyword.Trim()}";

        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: false, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> SearchFilesPageAsync(
        string keyword,
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        CancellationToken cancellationToken)
    {
        var query = string.IsNullOrWhiteSpace(keyword)
            ? "file:"
            : $"file: {keyword.Trim()}";

        return Task.Run(
            () => QueryPageInternal(query, ToEverythingSort(sortOption), offset, pageSize, isDirectoryResult: false, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> SearchFoldersAsync(
        string keyword,
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        var query = string.IsNullOrWhiteSpace(keyword)
            ? "folder:"
            : $"folder: {keyword.Trim()}";

        return Task.Run(
            () => QueryInternal(query, ToEverythingSort(sortOption), maxResults, isDirectoryResult: true, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> SearchFoldersPageAsync(
        string keyword,
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        CancellationToken cancellationToken)
    {
        var query = string.IsNullOrWhiteSpace(keyword)
            ? "folder:"
            : $"folder: {keyword.Trim()}";

        return Task.Run(
            () => QueryPageInternal(query, ToEverythingSort(sortOption), offset, pageSize, isDirectoryResult: true, CandidateSource.EverythingSearch, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFilesAsync(
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        return Task.Run(
            () => QueryInternal("file:", ToEverythingSort(sortOption), maxResults, isDirectoryResult: false, CandidateSource.EverythingRecent, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFilesPageAsync(
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        CancellationToken cancellationToken)
    {
        return Task.Run(
            () => QueryPageInternal("file:", ToEverythingSort(sortOption), offset, pageSize, isDirectoryResult: false, CandidateSource.EverythingRecent, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFoldersAsync(
        EverythingSortOption sortOption,
        int maxResults,
        CancellationToken cancellationToken)
    {
        return Task.Run(
            () => QueryInternal("folder:", ToEverythingSort(sortOption), maxResults, isDirectoryResult: true, CandidateSource.EverythingRecent, cancellationToken),
            cancellationToken);
    }

    public Task<IReadOnlyList<FileCandidate>> RecentFoldersPageAsync(
        EverythingSortOption sortOption,
        int offset,
        int pageSize,
        CancellationToken cancellationToken)
    {
        return Task.Run(
            () => QueryPageInternal("folder:", ToEverythingSort(sortOption), offset, pageSize, isDirectoryResult: true, CandidateSource.EverythingRecent, cancellationToken),
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
        var query = string.IsNullOrWhiteSpace(keyword)
            ? "file:"
            : $"file: {keyword.Trim()}";

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
        return Task.Run(
            () => QueryPageFilteredInternal(
                "file:",
                ToEverythingSort(sortOption),
                offset,
                pageSize,
                CandidateSource.EverythingRecent,
                allowedExtensions,
                cancellationToken),
            cancellationToken);
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
                Native.Everything_SetMatchPath(false);
                Native.Everything_SetMatchCase(false);
                Native.Everything_SetRegex(false);
                Native.Everything_SetRequestFlags(RequestFlags.FullPathAndFileName);
                Native.Everything_SetSort(sortType);
                Native.Everything_SetOffset((uint)boundedOffset);
                Native.Everything_SetMax((uint)boundedMaxResults);
                Native.Everything_SetSearchW(query);

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

                    var displayName = Path.GetFileName(fullPath);
                    if (string.IsNullOrWhiteSpace(displayName))
                    {
                        displayName = fullPath;
                    }

                    result.Add(
                        new FileCandidate
                        {
                            FullPath = fullPath,
                            DisplayName = displayName,
                            SecondaryText = Path.GetDirectoryName(fullPath) ?? string.Empty,
                            IsDirectory = isDirectoryResult,
                            Source = source,
                        });
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
                Native.Everything_SetMatchPath(false);
                Native.Everything_SetMatchCase(false);
                Native.Everything_SetRegex(false);
                Native.Everything_SetRequestFlags(RequestFlags.FullPathAndFileName);
                Native.Everything_SetSort(sortType);
                Native.Everything_SetOffset((uint)boundedOffset);
                Native.Everything_SetMax((uint)boundedMaxResults);
                Native.Everything_SetSearchW(query);

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

                    var displayName = Path.GetFileName(fullPath);
                    if (string.IsNullOrWhiteSpace(displayName))
                    {
                        displayName = fullPath;
                    }

                    result.Add(
                        new FileCandidate
                        {
                            FullPath = fullPath,
                            DisplayName = displayName,
                            SecondaryText = Path.GetDirectoryName(fullPath) ?? string.Empty,
                            IsDirectory = false,
                            Source = source,
                        });
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

        return allowedExtensions.Contains(extension.ToLowerInvariant());
    }

    private bool EnsureAvailableLocked()
    {
        if (!_isAvailable)
        {
            return false;
        }

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
