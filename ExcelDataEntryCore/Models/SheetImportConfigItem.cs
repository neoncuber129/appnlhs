using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class SheetImportConfigItem : ObservableObject
{
    private bool _isSelected;
    private int _headerFirstRow = 1;
    private int _headerLastRow = 3;
    private int _firstDataRow = 4;
    private int _sampleRow;
    private bool _isNameSheet;
    private int _nameColumnIndex = 2;

    public required string SheetName { get; init; }

    public bool IsHidden { get; init; }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }

    public int HeaderFirstRow
    {
        get => _headerFirstRow;
        set
        {
            if (!SetProperty(ref _headerFirstRow, value))
            {
                return;
            }

            if (HeaderLastRow < value)
            {
                HeaderLastRow = value;
            }

            RefreshColumnOptionsRequested?.Invoke(this, EventArgs.Empty);
        }
    }

    public int HeaderLastRow
    {
        get => _headerLastRow;
        set
        {
            if (!SetProperty(ref _headerLastRow, value))
            {
                return;
            }

            if (HeaderFirstRow > value)
            {
                HeaderFirstRow = value;
            }

            if (FirstDataRow <= value)
            {
                FirstDataRow = value + 1;
            }

            RefreshColumnOptionsRequested?.Invoke(this, EventArgs.Empty);
        }
    }

    public event EventHandler? RefreshColumnOptionsRequested;

    public int FirstDataRow
    {
        get => _firstDataRow;
        set => SetProperty(ref _firstDataRow, value);
    }

    public int SampleRow
    {
        get => _sampleRow;
        set => SetProperty(ref _sampleRow, value);
    }

    public bool IsNameSheet
    {
        get => _isNameSheet;
        set => SetProperty(ref _isNameSheet, value);
    }

    public int NameColumnIndex
    {
        get => _nameColumnIndex;
        set => SetProperty(ref _nameColumnIndex, value);
    }

    public List<HeaderColumnOption> ColumnOptions { get; } = [];

    public SheetImportConfig ToConfig() => new()
    {
        SheetName = SheetName,
        IsHidden = IsHidden,
        HeaderFirstRow = HeaderFirstRow,
        HeaderLastRow = HeaderLastRow,
        FirstDataRow = FirstDataRow,
        SampleRow = SampleRow,
        IsNameSheet = IsNameSheet,
        NameColumnIndex = NameColumnIndex
    };
}

public sealed class HeaderColumnOption
{
    public int ColumnIndex { get; init; }

    public string DisplayName { get; init; } = string.Empty;

    public override string ToString() => $"Cột {ColumnIndex}: {DisplayName.Replace('\n', ' ')}";
}
