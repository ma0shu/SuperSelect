namespace SuperSelect.App.Models;

internal enum OverlayMode
{
    Search,
    Recent,
    Tray,
    Explorer,
}

internal enum EverythingSortOption
{
    NameAsc,
    NameDesc,
    PathAsc,
    PathDesc,
    DateModifiedDesc,
    DateModifiedAsc,
}

internal enum CandidateSource
{
    EverythingSearch,
    EverythingRecent,
    Tray,
    Explorer,
}

internal sealed class FileCandidate
{
    public required string FullPath { get; init; }
    public required string DisplayName { get; init; }
    public string SecondaryText { get; init; } = string.Empty;
    public bool IsDirectory { get; init; }
    public CandidateSource Source { get; init; }
    public bool IsTrayPinned { get; init; }

    public string SourceLabel => Source switch
    {
        CandidateSource.EverythingSearch => "搜索",
        CandidateSource.EverythingRecent => "最近",
        CandidateSource.Tray => "托盘",
        CandidateSource.Explorer => "路径",
        _ => "未知",
    };
}
