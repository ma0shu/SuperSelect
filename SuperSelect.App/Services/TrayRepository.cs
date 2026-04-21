using System.IO;
using System.Text.Json;
using SuperSelect.App.Models;

namespace SuperSelect.App.Services;

internal sealed class TrayRepository
{
    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };

    private readonly object _syncRoot = new();
    private readonly string _storePath;
    private readonly List<string> _paths;
    private readonly HashSet<string> _pathIndex;

    public TrayRepository()
    {
        var storeDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SuperSelect");

        Directory.CreateDirectory(storeDirectory);
        _storePath = Path.Combine(storeDirectory, "tray.json");
        _paths = Load();
        _pathIndex = new HashSet<string>(_paths, StringComparer.OrdinalIgnoreCase);
    }

    public int Count
    {
        get
        {
            lock (_syncRoot)
            {
                return _paths.Count;
            }
        }
    }

    public int AddMany(IEnumerable<string> paths)
    {
        var added = 0;

        lock (_syncRoot)
        {
            foreach (var rawPath in paths)
            {
                var normalized = NormalizePath(rawPath);
                if (normalized is null)
                {
                    continue;
                }

                if (_pathIndex.Add(normalized))
                {
                    _paths.Add(normalized);
                    added++;
                }
            }

            if (added > 0)
            {
                SaveLocked();
            }
        }

        return added;
    }

    public bool Contains(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        lock (_syncRoot)
        {
            if (_pathIndex.Contains(path))
            {
                return true;
            }
        }

        var normalized = NormalizePath(path);
        if (normalized is null)
        {
            return false;
        }

        lock (_syncRoot)
        {
            return _pathIndex.Contains(normalized);
        }
    }

    public HashSet<string> GetPathIndexSnapshot()
    {
        lock (_syncRoot)
        {
            return new HashSet<string>(_pathIndex, StringComparer.OrdinalIgnoreCase);
        }
    }

    public bool Remove(string path)
    {
        var normalized = NormalizePath(path);
        if (normalized is null)
        {
            return false;
        }

        lock (_syncRoot)
        {
            if (!_pathIndex.Remove(normalized))
            {
                return false;
            }

            var idx = _paths.FindIndex(p => string.Equals(p, normalized, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
            {
                _paths.RemoveAt(idx);
                SaveLocked();
                return true;
            }

            _pathIndex.Add(normalized);
            return false;
        }
    }

    public void PinToTop(string path)
    {
        var normalized = NormalizePath(path);
        if (normalized is null)
        {
            return;
        }

        lock (_syncRoot)
        {
            var idx = _paths.FindIndex(p => string.Equals(p, normalized, StringComparison.OrdinalIgnoreCase));
            if (idx > 0)
            {
                _paths.RemoveAt(idx);
                _paths.Insert(0, normalized);
                SaveLocked();
            }
        }
    }

    public void Clear()
    {
        lock (_syncRoot)
        {
            _paths.Clear();
            _pathIndex.Clear();
            SaveLocked();
        }
    }

    public IReadOnlyList<FileCandidate> Query(string keyword)
    {
        var filter = keyword.Trim();
        List<string> snapshot;

        lock (_syncRoot)
        {
            snapshot = new List<string>(_paths);
        }

        if (snapshot.Count == 0)
        {
            return [];
        }

        if (string.IsNullOrWhiteSpace(filter))
        {
            return BuildCandidates(snapshot);
        }

        var matched = new List<string>(snapshot.Count);
        for (var i = 0; i < snapshot.Count; i++)
        {
            var path = snapshot[i];
            if (PathMatches(path, filter))
            {
                matched.Add(path);
            }
        }

        if (matched.Count == 0)
        {
            return [];
        }

        matched.Sort(
            static (left, right) =>
                StringComparer.OrdinalIgnoreCase.Compare(Path.GetFileName(left), Path.GetFileName(right)));
        return BuildCandidates(matched);
    }

    private List<string> Load()
    {
        if (!File.Exists(_storePath))
        {
            return new List<string>();
        }

        try
        {
            var json = File.ReadAllText(_storePath);
            var paths = JsonSerializer.Deserialize<List<string>>(json) ?? [];
            var result = new List<string>(paths.Count);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < paths.Count; i++)
            {
                var normalized = NormalizePath(paths[i]);
                if (normalized is null || !seen.Add(normalized))
                {
                    continue;
                }

                result.Add(normalized);
            }

            return result;
        }
        catch
        {
            return new List<string>();
        }
    }

    private void SaveLocked()
    {
        var payload = JsonSerializer.Serialize(_paths, WriteOptions);
        File.WriteAllText(_storePath, payload);
    }

    private static IReadOnlyList<FileCandidate> BuildCandidates(IReadOnlyList<string> paths)
    {
        var result = new List<FileCandidate>(paths.Count);
        for (var i = 0; i < paths.Count; i++)
        {
            var path = paths[i];
            var fileName = Path.GetFileName(path);
            result.Add(
                new FileCandidate
                {
                    FullPath = path,
                    DisplayName = string.IsNullOrWhiteSpace(fileName) ? path : fileName,
                    SecondaryText = Path.GetDirectoryName(path) ?? string.Empty,
                    IsDirectory = Directory.Exists(path),
                    Source = CandidateSource.Tray,
                    IsTrayPinned = true,
                });
        }

        return result;
    }

    private static bool PathMatches(string path, string filter)
    {
        return path.Contains(filter, StringComparison.OrdinalIgnoreCase)
            || Path.GetFileName(path).Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private static string? NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            var trimmed = path.Trim().Trim('"');
            return Path.GetFullPath(trimmed);
        }
        catch
        {
            return null;
        }
    }
}
