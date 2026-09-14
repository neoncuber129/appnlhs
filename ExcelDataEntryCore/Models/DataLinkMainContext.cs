namespace ExcelDataEntryApp.Models;

public sealed class DataLinkMainContext
{
    public string KeySheetName { get; init; } = string.Empty;

    public int KeySheetFirstDataRow { get; init; }

    public IReadOnlyDictionary<string, int> FirstDataRowBySheet { get; init; }
        = new Dictionary<string, int>(StringComparer.Ordinal);

    public IReadOnlyDictionary<string, int> FormatReferenceRowBySheet { get; init; }
        = new Dictionary<string, int>(StringComparer.Ordinal);

    public bool IsMultiSheet => FirstDataRowBySheet.Count > 0;
}
