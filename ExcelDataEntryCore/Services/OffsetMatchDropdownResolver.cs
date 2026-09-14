using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ExcelDataEntryApp.Services;

/// <summary>
/// Giải công thức dropdown phụ thuộc dạng OFFSET/MATCH/COUNTIF (x14:dataValidation).
/// </summary>
public static class OffsetMatchDropdownResolver
{
    private static readonly Regex OffsetMatchCountifRegex = new(
        @"OFFSET\(\s*'(?<catalogSheet>[^']+)'\s*!\s*\$(?<anchorCol>[A-Za-z]+)\s*\$(?<anchorRow>\d+)\s*,\s*"
        + @"MATCH\(\s*(?<parentCol>[A-Za-z]+)(?<parentRow>\d+)\s*,\s*'(?<matchSheet>[^']+)'\s*!\s*\$(?<matchCol>[A-Za-z]+)\s*:\s*\$(?<matchColEnd>[A-Za-z]+)\s*,\s*0\s*\)\s*-\s*(?<rowAdjust>\d+)\s*,\s*"
        + @"(?<colOffset>\d+)\s*,\s*COUNTIF\(\s*'(?<countSheet>[^']+)'\s*!\s*\$(?<countCol>[A-Za-z]+)\s*:\s*\$(?<countColEnd>[A-Za-z]+)\s*,\s*(?<parentCol2>[A-Za-z]+)(?<parentRow2>\d+)\s*\)\s*,\s*1\s*\)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static bool TryResolve(
        ExcelPackage package,
        int dataRowIndex,
        int dataColumnIndex,
        ExtendedListValidationRule rule,
        Func<int, int, string?> getCellValue,
        out IReadOnlyList<string> values)
    {
        values = [];
        var cellRange = rule.FindRange(dataRowIndex, dataColumnIndex);
        if (cellRange is null)
        {
            return false;
        }

        var normalized = rule.Formula.Trim().TrimStart('=');
        var match = OffsetMatchCountifRegex.Match(normalized);
        if (!match.Success)
        {
            return false;
        }

        if (!int.TryParse(match.Groups["anchorRow"].Value, out var anchorRow)
            || !int.TryParse(match.Groups["rowAdjust"].Value, out var rowAdjust)
            || !int.TryParse(match.Groups["colOffset"].Value, out var colOffset)
            || !int.TryParse(match.Groups["parentRow"].Value, out var parentTemplateRow))
        {
            return false;
        }

        var parentColumn = ExcelCellAddress.ColumnLettersToIndex(match.Groups["parentCol"].Value);
        var anchorColumn = ExcelCellAddress.ColumnLettersToIndex(match.Groups["anchorCol"].Value);
        var matchColumn = ExcelCellAddress.ColumnLettersToIndex(match.Groups["matchCol"].Value);
        var catalogSheetName = match.Groups["catalogSheet"].Value;

        var rowDelta = dataRowIndex - cellRange.Value.StartRow;
        var parentRow = parentTemplateRow + rowDelta;
        var parentValue = getCellValue(parentRow, parentColumn)?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(parentValue))
        {
            return true;
        }

        var catalogWorksheet = package.Workbook.Worksheets
            .FirstOrDefault(ws => string.Equals(ws.Name, catalogSheetName, StringComparison.Ordinal));
        if (catalogWorksheet is null)
        {
            return false;
        }

        var endRow = catalogWorksheet.Dimension?.End.Row ?? 0;
        var matchRow = -1;
        var matchCount = 0;
        for (var row = 1; row <= endRow; row++)
        {
            var text = catalogWorksheet.Cells[row, matchColumn].Text?.Trim() ?? string.Empty;
            if (!string.Equals(text, parentValue, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (matchRow < 0)
            {
                matchRow = row;
            }

            matchCount++;
        }

        if (matchRow < 0 || matchCount <= 0)
        {
            return true;
        }

        var rowOffset = matchRow - rowAdjust;
        var resultStartRow = anchorRow + rowOffset;
        var resultColumn = anchorColumn + colOffset;
        var results = new List<string>();
        for (var row = resultStartRow; row < resultStartRow + matchCount; row++)
        {
            if (row < 1)
            {
                continue;
            }

            var text = catalogWorksheet.Cells[row, resultColumn].Text?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text))
            {
                results.Add(text);
            }
        }

        values = results
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return true;
    }
}
