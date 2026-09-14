using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public enum ImportListMatchStatus
{
    Ready
}

public sealed class ImportListRowOption : ObservableObject
{
    private bool _isSelected;

    public required int DataRowIndex { get; init; }

    public required string KeyValue { get; init; }

    public required string NormalizedKey { get; init; }

    public IReadOnlyDictionary<int, string> CellValuesByColumn { get; init; }
        = new Dictionary<int, string>();

    public string RowSearchText { get; init; } = string.Empty;

    public ImportListMatchStatus Status { get; init; }

    public bool CanSelect => Status == ImportListMatchStatus.Ready;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (!CanSelect)
            {
                return;
            }

            SetProperty(ref _isSelected, value);
        }
    }

    public string Display => $"{KeyValue} (dòng {DataRowIndex})";

    public string GetCellText(int columnIndex) =>
        CellValuesByColumn.TryGetValue(columnIndex, out var value) ? value : string.Empty;

    public string this[int columnIndex] => GetCellText(columnIndex);
}

public sealed class ImportListLinkRequest
{
    public required ImportListProfile Profile { get; init; }

    public required IReadOnlyList<int> SelectedDataRowIndices { get; init; }
}
