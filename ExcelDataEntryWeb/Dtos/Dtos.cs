using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;

namespace ExcelDataEntryWeb.Dtos;

// ─── Request DTOs ─────────────────────────────────────────────────────────────

public class OpenWorkbookPathRequest
{
    public string Path { get; set; } = string.Empty;
}

public class SelectSheetRequest
{
    public string SheetName { get; set; } = string.Empty;
    public int HeaderRow { get; set; } = 3;
    public int NameColumn { get; set; } = 2;
    public int SampleRow { get; set; } = 4;
    public bool AutoSkipBlank { get; set; } = true;
    public bool SuggestionsDisabled { get; set; } = false;
    public bool AutoSave { get; set; } = true;
}

public class SaveAsRequest
{
    public string Path { get; set; } = string.Empty;
}

public class AddRecordRequest
{
    public string? Name { get; set; }
}

public class RenameRecordRequest
{
    public int LogicalIndex { get; set; }
    public string NewName { get; set; } = string.Empty;
}

public class UpdateCellRequest
{
    public string? SheetName { get; set; }
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public string? Value { get; set; }
}

public class SetVisibilityRequest
{
    public bool IsVisible { get; set; }
}

public class ShowAllRequest
{
    public bool Show { get; set; }
}

public class SortRequest
{
    public int ColumnIndex { get; set; }
    public ColumnSortMode Mode { get; set; }
    public IReadOnlyList<string>? CustomValues { get; set; }
}

public class HsskExportRequest
{
    public string? TemplateDocxPath { get; set; }
    public int GenderColumnIndex { get; set; } = -1;
    public IReadOnlyList<int>? SkipSampleColumnIndexes { get; set; }
}

public class ImportFileExportRequest
{
    public string? TemplatePath { get; set; }
    public int GenderColumnIndex { get; set; } = -1;
    public IReadOnlyList<int>? SkipSampleColumnIndexes { get; set; }
}

public class DependentDropdownRequest
{
    public string? SheetName { get; set; }
    public int RowIndex { get; set; }
    public List<CellOverrideDto>? Overrides { get; set; }
}

public class CellOverrideDto
{
    public string SheetName { get; set; } = string.Empty;
    public int RowIndex { get; set; }
    public int ColumnIndex { get; set; }
    public string Value { get; set; } = string.Empty;
}

// ─── Response DTOs ────────────────────────────────────────────────────────────

public record WorkbookInfoDto(
    bool IsLoaded, string FilePath, string SelectedSheet, string StatusMessage,
    bool IsMultiSheetMode, List<string> Sheets,
    int HeaderRowNumber, int NameColumnIndex, int SampleRowThreshold,
    bool IsAutoSaveEnabled, bool SuggestionsDisabled, bool AutoSkipBlankHeaders);

public record RecordDto(int RowIndex, int LogicalIndex, string KeyDisplay, string TooltipPreview);

public record HeaderDto(int ColumnIndex, string Name, bool IsVisible, string SheetName, string ToggleLabel);

public class EditableFieldDto
{
    public string HeaderName { get; init; } = string.Empty;
    public string HeaderDisplayName { get; init; } = string.Empty;
    public string ParentHeaderName { get; init; } = string.Empty;
    public bool ShowParentHeader { get; init; }
    public bool IsGroupedUnderParentHeader { get; init; }
    public bool IsFirstInHeaderGroup { get; init; }
    public bool IsLastInHeaderGroup { get; init; }
    public string GroupBorderHex { get; init; } = "Transparent";
    public string ParentTitleBackgroundHex { get; init; } = "Transparent";
    public string ParentTitleForegroundHex { get; init; } = "#111827";
    public int ColumnIndex { get; init; }
    public string SheetName { get; init; } = string.Empty;
    public bool ShowSheetSeparator { get; init; }
    public string SheetSeparatorTitle { get; init; } = string.Empty;
    public int SourceRowIndex { get; init; }
    public string Value { get; init; } = string.Empty;
    public bool HasDropdown { get; init; }
    public bool HasLargeDropdown { get; init; }
    public bool HasSmallDropdown { get; init; }
    public bool ShowDropdownEditor { get; init; }
    public bool IsDependentDropdown { get; init; }
    public IReadOnlyList<int> ParentDropdownColumns { get; init; } = [];
    public IReadOnlyList<string> DropdownOptions { get; init; } = [];
    public IReadOnlyList<string> SuggestionOptions { get; init; } = [];
    public bool HasSuggestions { get; init; }
    public bool IsDropdownValueInvalid { get; init; }
    public string DropdownValidationMessage { get; init; } = string.Empty;
    public string RowHighlightBackgroundHex { get; init; } = "Transparent";
    public string RowHighlightBorderHex { get; init; } = "Transparent";
    public string RowHeaderForegroundHex { get; init; } = "#111827";
    public string RowInputBackgroundHex { get; init; } = "#FFFFFF";
    public string RowInputForegroundHex { get; init; } = "#111827";
    public bool HasRowHighlight { get; init; }

    public static EditableFieldDto From(EditableField f) => new()
    {
        HeaderName = f.HeaderName,
        HeaderDisplayName = f.HeaderDisplayName,
        ParentHeaderName = f.ParentHeaderName,
        ShowParentHeader = f.ShowParentHeader,
        IsGroupedUnderParentHeader = f.IsGroupedUnderParentHeader,
        IsFirstInHeaderGroup = f.IsFirstInHeaderGroup,
        IsLastInHeaderGroup = f.IsLastInHeaderGroup,
        GroupBorderHex = f.GroupBorderHex,
        ParentTitleBackgroundHex = f.ParentTitleBackgroundHex,
        ParentTitleForegroundHex = f.ParentTitleForegroundHex,
        ColumnIndex = f.ColumnIndex,
        SheetName = f.SheetName,
        ShowSheetSeparator = f.ShowSheetSeparator,
        SheetSeparatorTitle = f.SheetSeparatorTitle,
        SourceRowIndex = f.SourceRowIndex,
        Value = f.Value,
        HasDropdown = f.HasDropdown,
        HasLargeDropdown = f.HasLargeDropdown,
        HasSmallDropdown = f.HasSmallDropdown,
        ShowDropdownEditor = f.ShowDropdownEditor,
        IsDependentDropdown = f.IsDependentDropdown,
        ParentDropdownColumns = f.ParentDropdownColumns,
        DropdownOptions = f.DropdownOptions.Count <= 50 ? f.DropdownOptions : f.DisplayDropdownItems,
        SuggestionOptions = f.SuggestionOptions,
        HasSuggestions = f.HasSuggestions,
        IsDropdownValueInvalid = f.IsDropdownValueInvalid,
        DropdownValidationMessage = f.DropdownValidationMessage,
        RowHighlightBackgroundHex = f.RowHighlightBackgroundHex,
        RowHighlightBorderHex = f.RowHighlightBorderHex,
        RowHeaderForegroundHex = f.RowHeaderForegroundHex,
        RowInputBackgroundHex = f.RowInputBackgroundHex,
        RowInputForegroundHex = f.RowInputForegroundHex,
        HasRowHighlight = f.HasRowHighlight
    };
}
