using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class HsskFemaleColumnOption : ObservableObject
{
    private bool _isSelected;

    public required int ColumnIndex { get; init; }

    public required string HeaderName { get; init; }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }

    public string Display => $"${ColumnIndex} — {HeaderName}";
}
