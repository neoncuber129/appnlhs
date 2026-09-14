namespace ExcelDataEntryApp.Models;

public sealed class SortCustomValueItem
{
    public required string Value { get; init; }

    public string Display => string.IsNullOrWhiteSpace(Value) ? "(trống)" : Value;
}
