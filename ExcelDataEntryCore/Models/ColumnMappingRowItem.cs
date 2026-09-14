using ExcelDataEntryApp.Infrastructure;

namespace ExcelDataEntryApp.Models;

public sealed class ColumnMappingRowItem : ObservableObject
{
    private HeaderDefinition? _targetColumn;
    private HeaderDefinition? _sourceColumn;

    public HeaderDefinition? TargetColumn
    {
        get => _targetColumn;
        set => SetProperty(ref _targetColumn, value);
    }

    public HeaderDefinition? SourceColumn
    {
        get => _sourceColumn;
        set => SetProperty(ref _sourceColumn, value);
    }
}
