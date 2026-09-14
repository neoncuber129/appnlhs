using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class HeaderDefinition : ObservableObject
{
    private bool _isVisible = true;

    public required string Name { get; init; }

    public required int ColumnIndex { get; init; }

    /// <summary>Sheet chứa cột (chế độ ghép nhiều sheet). Rỗng khi nhập 1 sheet.</summary>
    public string SheetName { get; init; } = string.Empty;

    public int HeaderFirstRow { get; init; } = 1;

    public int HeaderLastRow { get; init; }

    public bool IsVisible
    {
        get => _isVisible;
        set
        {
            if (SetProperty(ref _isVisible, value))
            {
                OnPropertyChanged(nameof(ToggleLabel));
            }
        }
    }

    public string ToggleLabel => IsVisible ? "Ẩn" : "Hiện";

    public string LinkDisplayName => string.IsNullOrEmpty(SheetName)
        ? $"Cột {ColumnIndex}: {Name.Replace('\n', ' ')}"
        : $"[{SheetName}] Cột {ColumnIndex}: {Name.Replace('\n', ' ')}";
}
