namespace ExcelDataEntryApp.Models;

public sealed class ImportFileSheetPlan
{
    public required string SheetName { get; init; }

    public int FirstDataRow { get; init; }

    public int LastDataRow { get; init; }

    public int SampleRow { get; init; }

    public IReadOnlyList<int> ColumnIndexes { get; init; } = [];

    public IReadOnlyList<int> SkipSampleColumnIndexes { get; init; } = [];

    public int GenderColumnIndex { get; init; }

    public string? GenderSheetName { get; init; }

    public int GenderSheetFirstDataRow { get; init; }
}

public sealed class ImportFileExportPreview
{
    public bool IsMultiSheetMode { get; init; }

    public string NameSheetName { get; init; } = string.Empty;

    public int FirstDataRow { get; init; }

    public int LastNameRow { get; init; }

    public int SampleRow { get; init; }

    public int SheetCount { get; init; }

    public int RowCount => LastNameRow >= FirstDataRow ? LastNameRow - FirstDataRow + 1 : 0;
}

public sealed class ImportFileExportResult
{
    public int FilledCellCount { get; init; }

    public int ProcessedRowCount { get; init; }

    public int ProcessedSheetCount { get; init; }

    public string OutputPath { get; init; } = string.Empty;
}
