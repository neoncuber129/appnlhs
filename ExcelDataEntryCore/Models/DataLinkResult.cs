namespace ExcelDataEntryApp.Models;

public sealed class DuplicateKeyEntry
{
    public required string KeyValue { get; init; }
    public int OccurrenceCount { get; init; }
    public required IReadOnlyList<int> RowNumbers { get; init; }
    public IReadOnlyList<string> SpellingVariants { get; init; } = [];
}

public enum SkippedLinkReason
{
    MainDuplicateKey,
    DataDuplicateKey,
    NoMatchInDataFile
}

public sealed class SkippedLinkEntry
{
    public int RowIndex { get; init; }
    public required string KeyValue { get; init; }
    public SkippedLinkReason Reason { get; init; }
}

public sealed class DataLinkResult
{
    public int RowsMatched { get; init; }

    public int RowsCreatedInMain { get; init; }

    public int CellsWritten { get; init; }

    public int SkippedSelectedNotInMain { get; init; }
    public IReadOnlyList<DuplicateKeyEntry> MainFileDuplicates { get; init; } = [];
    public IReadOnlyList<DuplicateKeyEntry> DataFileDuplicates { get; init; } = [];
    public IReadOnlyList<string> KeysOnlyInMain { get; init; } = [];
    public IReadOnlyList<string> KeysOnlyInData { get; init; } = [];
    public IReadOnlyList<SkippedLinkEntry> SkippedRows { get; init; } = [];

    public bool HasNotableIssues =>
        SkippedRows.Count > 0
        || MainFileDuplicates.Count > 0
        || DataFileDuplicates.Count > 0
        || KeysOnlyInMain.Count > 0
        || KeysOnlyInData.Count > 0;
}

public sealed class DataLinkAnalysis
{
    public IReadOnlyList<DuplicateKeyEntry> MainFileDuplicates { get; init; } = [];
    public IReadOnlyList<DuplicateKeyEntry> DataFileDuplicates { get; init; } = [];
    public IReadOnlyList<string> KeysOnlyInMain { get; init; } = [];
    public IReadOnlyList<string> KeysOnlyInData { get; init; } = [];
}
