using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class ImportListColumnFilterState : ObservableObject
{
    private HashSet<string>? _allowedValues;

    public required int ColumnIndex { get; init; }

    public required string HeaderName { get; set; }

    /// <summary>null = không lọc (hiện tất cả giá trị).</summary>
    public HashSet<string>? AllowedValues
    {
        get => _allowedValues;
        set
        {
            if (!SetProperty(ref _allowedValues, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsActive));
            OnPropertyChanged(nameof(FilterGlyph));
        }
    }

    public bool IsActive => AllowedValues is not null;

    public string FilterGlyph => IsActive ? "▾*" : "▾";
}
