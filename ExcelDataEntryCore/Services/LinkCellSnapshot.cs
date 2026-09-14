namespace ExcelDataEntryApp.Services;

/// <summary>Giá trị ô file số liệu: chuỗi hiển thị (Text) và giá trị gốc (Value).</summary>
public sealed class LinkCellSnapshot
{
    public string DisplayText { get; init; } = string.Empty;
    public object? RawValue { get; init; }
}
