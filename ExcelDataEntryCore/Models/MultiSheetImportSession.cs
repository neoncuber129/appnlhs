namespace ExcelDataEntryApp.Models;

public sealed class MultiSheetImportSession
{
    public IReadOnlyList<SheetImportConfig> Sheets { get; init; } = [];

    public string NameSheetName { get; init; } = string.Empty;

    public int NameColumnIndex { get; init; } = MultiSheetImportDefaults.NameColumnIndex;

    public bool ShowHiddenSheets { get; init; }

    public bool AutoSkipBlankHeaders { get; init; } = MultiSheetImportDefaults.AutoSkipBlankHeaders;

    public bool SuggestionsDisabled { get; init; } = MultiSheetImportDefaults.SuggestionsDisabled;

    public bool IsActive => Sheets.Count > 0;

    public SheetImportConfig? NameSheet =>
        Sheets.FirstOrDefault(s => string.Equals(s.SheetName, NameSheetName, StringComparison.Ordinal));

    public SheetImportConfig? GetSheetConfig(string sheetName) =>
        Sheets.FirstOrDefault(s => string.Equals(s.SheetName, sheetName, StringComparison.Ordinal));

    public int GetDataRowForLogicalIndex(string sheetName, int logicalIndex)
    {
        var cfg = GetSheetConfig(sheetName);
        return cfg is null ? -1 : cfg.FirstDataRow + logicalIndex;
    }
}
