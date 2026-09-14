using OfficeOpenXml;

namespace ExcelDataEntryApp.Services;

public sealed class HeaderMergeResolver
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<MergeRange> _merges;

    public HeaderMergeResolver(ExcelWorksheet worksheet)
    {
        _worksheet = worksheet;
        _merges = ParseMerges(worksheet);
    }

    public string? GetEffectiveText(int row, int column)
    {
        if (row < 1 || column < 1)
        {
            return null;
        }

        var anchor = FindMergeAnchor(row, column);
        return _worksheet.Cells[anchor.Row, anchor.Column].Text;
    }

    public IReadOnlyList<string> BuildColumnHeaderParts(int column, int firstRow, int lastRow)
    {
        var parts = new List<string>();
        string? lastPart = null;

        for (var row = firstRow; row <= lastRow; row++)
        {
            var text = GetEffectiveText(row, column)?.Trim();
            if (string.IsNullOrWhiteSpace(text) || string.Equals(text, lastPart, StringComparison.Ordinal))
            {
                continue;
            }

            parts.Add(text);
            lastPart = text;
        }

        return parts;
    }

    private (int Row, int Column) FindMergeAnchor(int row, int column)
    {
        foreach (var merge in _merges)
        {
            if (row >= merge.FromRow && row <= merge.ToRow
                && column >= merge.FromCol && column <= merge.ToCol)
            {
                return (merge.FromRow, merge.FromCol);
            }
        }

        return (row, column);
    }

    private static List<MergeRange> ParseMerges(ExcelWorksheet worksheet)
    {
        var merges = new List<MergeRange>();
        foreach (var address in worksheet.MergedCells)
        {
            var range = worksheet.Cells[address];
            merges.Add(new MergeRange(
                range.Start.Row,
                range.Start.Column,
                range.End.Row,
                range.End.Column));
        }

        return merges;
    }

    private readonly record struct MergeRange(int FromRow, int FromCol, int ToRow, int ToCol);
}
