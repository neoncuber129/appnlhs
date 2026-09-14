namespace ExcelDataEntryApp.Models;

public sealed class RecordItem
{
    /// <summary>Chỉ số người thứ N (0-based) khi ghép nhiều sheet; -1 khi nhập 1 sheet.</summary>
    public int LogicalIndex { get; init; } = -1;

    public required int RowIndex { get; init; }

    public required string KeyDisplay { get; set; }

    public required string TooltipPreview { get; init; }
}
