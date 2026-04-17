using System.IO;
using System.Text.Json;
using SuperSelect.App.Models;

namespace SuperSelect.App.Services;

internal sealed class TrayRepository
{
    private readonly object _syncRoot = new();
    private readonly string _storePath;
    private readonly HashSet<string> _paths;

    public TrayRepository()
    {
        var storeDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SuperSelect");

        Directory.CreateDirectory(storeDirectory);
        _storePath = Path.Combine(storeDirectory, "tray.json");
        _paths = Load();
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

                if (_paths.Add(normalized))
                {
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

    public bool Remove(string path)
    {
        var normalized = NormalizePath(path);
        if (normalized is null)
        {
            return false;
        }

        lock (_syncRoot)
        {
            var removed = _paths.Remove(normalized);
            if (removed)
            {
                SaveLocked();
            }

            return removed;
        }
    }

    public void Clear()
    {
        lock (_syncRoot)
        {
            _paths.Clear();
            SaveLocked();
        }
    }

    public IReadOnlyList<FileCandidate> Query(string keyword)
    {
        var filter = keyword.Trim();

        lock (_syncRoot)
        {
            return _paths
                .Where(path => string.IsNullOrWhiteSpace(filter) || PathMatches(path, filter))
                .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
                .Select(
                    path =>
                    {
                        var fileName = Path.GetFileName(path);
                        return new FileCandidate
                        {
                            FullPath = path,
                            DisplayName = string.IsNullOrWhiteSpace(fileName) ? path : fileName,
                            SecondaryText = Path.GetDirectoryName(path) ?? string.Empty,
                            IsDirectory = Directory.Exists(path),
                            Source = CandidateSource.Tray,
                        };
                    })
                .ToList();
        }
    }

    private HashSet<string> Load()
    {
        if (!File.Exists(_storePath))
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        try
        {
            var json = File.ReadAllText(_storePath);
            var paths = JsonSerializer.Deserialize<List<string>>(json) ?? [];
            var normalized = paths
                .Select(NormalizePath)
                .Where(path => path is not null)!
                .Cast<string>();

            return new HashSet<string>(normalized, StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private void SaveLocked()
    {
        var payload = JsonSerializer.Serialize(
            _paths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase),
            new JsonSerializerOptions { WriteIndented = true });

        File.WriteAllText(_storePath, payload);
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
