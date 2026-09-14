using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class ExportRecordOption : ObservableObject
{
    private bool _isSelected;

    public required int RowIndex { get; init; }

    public required string Name { get; init; }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }

    public string Display => $"{Name} (dòng {RowIndex})";
}
