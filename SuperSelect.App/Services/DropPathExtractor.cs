using System.Collections.Specialized;
using System.Windows;
using IDataObject = System.Windows.IDataObject;

namespace SuperSelect.App.Services;

internal static class DropPathExtractor
{
    private static readonly IReadOnlyList<string> EmptyPaths = Array.Empty<string>();
    private static readonly string[] PreferredFormats = [System.Windows.DataFormats.FileDrop, "FileNameW", "FileName"];

    public static bool HasDroppedFiles(IDataObject data)
    {
        foreach (var format in PreferredFormats)
        {
            if (!data.GetDataPresent(format))
            {
                continue;
            }

            if (TryGetAnyPath(data.GetData(format)))
            {
                return true;
            }
        }

        foreach (var format in data.GetFormats())
        {
            object? raw;
            try
            {
                raw = data.GetData(format);
            }
            catch
            {
                continue;
            }

            if (TryGetAnyPath(raw))
            {
                return true;
            }
        }

        return false;
    }

    public static IReadOnlyList<string> ExtractDroppedPaths(IDataObject data)
    {
        foreach (var format in PreferredFormats)
        {
            if (!data.GetDataPresent(format))
            {
                continue;
            }

            IReadOnlyList<string> paths;
            try
            {
                paths = ExtractDroppedPathsFromRawData(data.GetData(format));
            }
            catch
            {
                continue;
            }

            if (paths.Count > 0)
            {
                return paths;
            }
        }

        foreach (var format in data.GetFormats())
        {
            IReadOnlyList<string> paths;
            try
            {
                paths = ExtractDroppedPathsFromRawData(data.GetData(format));
            }
            catch
            {
                continue;
            }

            if (paths.Count > 0)
            {
                return paths;
            }
        }

        return EmptyPaths;
    }

    private static bool TryGetAnyPath(object? raw)
    {
        if (raw is null)
        {
            return false;
        }

        if (raw is string single)
        {
            return !string.IsNullOrWhiteSpace(single);
        }

        if (raw is string[] paths)
        {
            for (var i = 0; i < paths.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(paths[i]))
                {
                    return true;
                }
            }

            return false;
        }

        if (raw is StringCollection collection)
        {
            for (var i = 0; i < collection.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(collection[i]))
                {
                    return true;
                }
            }

            return false;
        }

        if (raw is IEnumerable<string> enumerable)
        {
            foreach (var path in enumerable)
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    return true;
                }
            }

            return false;
        }

        return false;
    }

    private static IReadOnlyList<string> ExtractDroppedPathsFromRawData(object? raw)
    {
        if (raw is null)
        {
            return EmptyPaths;
        }

        if (raw is string single)
        {
            return string.IsNullOrWhiteSpace(single)
                ? EmptyPaths
                : [single];
        }

        if (raw is string[] pathArray)
        {
            return Deduplicate(pathArray);
        }

        if (raw is StringCollection collection && collection.Count > 0)
        {
            var values = new string[collection.Count];
            collection.CopyTo(values, 0);
            return Deduplicate(values);
        }

        if (raw is IEnumerable<string> enumerable)
        {
            return Deduplicate(enumerable);
        }

        return EmptyPaths;
    }

    private static IReadOnlyList<string> Deduplicate(IEnumerable<string> input)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();

        foreach (var path in input)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }

            if (seen.Add(path))
            {
                result.Add(path);
            }
        }

        return result.Count == 0 ? EmptyPaths : result;
    }
}
