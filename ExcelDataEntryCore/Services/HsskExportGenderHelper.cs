namespace ExcelDataEntryApp.Services;

public static class HsskExportGenderHelper
{
    public const string FemaleColumnMarker = "s";

    public static bool IsFemaleColumnMarker(string? cellText)
    {
        if (string.IsNullOrWhiteSpace(cellText))
        {
            return false;
        }

        return string.Equals(cellText.Trim(), FemaleColumnMarker, StringComparison.OrdinalIgnoreCase);
    }

    public static bool ShouldUseFemaleOnlyColumns(string? genderText)
    {
        if (string.IsNullOrWhiteSpace(genderText))
        {
            return false;
        }

        var normalized = genderText.Trim();
        if (normalized.Contains('ữ', StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("nu", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("female", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains('♀', StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (normalized.Contains("nam", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("male", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains('♂', StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return false;
    }
}
