using System.Text.RegularExpressions;

namespace ExcelDataEntryApp.Services;

public static class ExcelCellAddress
{
    private static readonly Regex CellRefRegex = new(
        @"^(\$?)([A-Za-z]+)(\$?)(\d+)$",
        RegexOptions.CultureInvariant);

    public static bool TryParseCell(string cellAddress, out int row, out int column)
    {
        row = 0;
        column = 0;
        var match = CellRefRegex.Match(cellAddress.Replace("$", string.Empty, StringComparison.Ordinal));
        if (!match.Success)
        {
            return false;
        }

        column = ColumnLettersToIndex(match.Groups[2].Value);
        return int.TryParse(match.Groups[4].Value, out row);
    }

    public static int ColumnLettersToIndex(string letters)
    {
        var result = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            result = (result * 26) + (ch - 'A' + 1);
        }

        return result;
    }

    public static string ColumnIndexToLetters(int columnIndex)
    {
        if (columnIndex < 1)
        {
            return string.Empty;
        }

        var letters = string.Empty;
        var value = columnIndex;
        while (value > 0)
        {
            value--;
            letters = (char)('A' + (value % 26)) + letters;
            value /= 26;
        }

        return letters;
    }
}
