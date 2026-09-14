using System.Text.RegularExpressions;

namespace ExcelDataEntryApp.Services;

/// <summary>
/// Dịch tham chiếu dòng trong công thức validation giống Excel khi áp dụng xuống từng dòng.
/// Chỉ dịch phần row không có $ (tương đối); bỏ qua ô đã gắn sheet (catalog).
/// </summary>
public static class ValidationFormulaAdjuster
{
    private static readonly Regex CellRefRegex = new(
        @"(?<![A-Za-z0-9_!'$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])",
        RegexOptions.CultureInvariant);

    public static string AdjustForRow(string formula, int validationStartRow, int currentRow)
    {
        var rowDelta = currentRow - validationStartRow;
        if (rowDelta == 0 || string.IsNullOrWhiteSpace(formula))
        {
            return formula;
        }

        return CellRefRegex.Replace(
            formula,
            match =>
            {
                if (match.Groups[3].Value == "$")
                {
                    return match.Value;
                }

                if (!int.TryParse(match.Groups[4].Value, out var row))
                {
                    return match.Value;
                }

                var adjustedRow = row + rowDelta;
                if (adjustedRow < 1)
                {
                    return match.Value;
                }

                return $"{match.Groups[1].Value}{match.Groups[2].Value}{match.Groups[3].Value}{adjustedRow}";
            });
    }

    public static bool LooksLikeDynamicFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return false;
        }

        var normalized = formula.Trim().TrimStart('=');
        return normalized.Contains("OFFSET(", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("MATCH(", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("INDIRECT(", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("INDEX(", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("COUNTIF(", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("FILTER(", StringComparison.OrdinalIgnoreCase);
    }

    public static IReadOnlyList<int> ExtractSameSheetParentColumns(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return [];
        }

        var withoutQualifiedRefs = Regex.Replace(
            formula,
            @"(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?",
            string.Empty,
            RegexOptions.CultureInvariant);

        var columns = new HashSet<int>();
        foreach (Match match in CellRefRegex.Matches(withoutQualifiedRefs))
        {
            var column = ExcelCellAddress.ColumnLettersToIndex(match.Groups[2].Value);
            if (column > 0)
            {
                columns.Add(column);
            }
        }

        return columns.OrderBy(c => c).ToList();
    }
}
