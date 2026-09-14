using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;

namespace ExcelDataEntryWeb;

/// <summary>
/// In-memory session giữ trạng thái workbook và nhập liệu cho một phiên web.
/// Tương đương MainViewModel nhưng không phụ thuộc WPF.
/// </summary>
public sealed class AppSession : IDisposable
{
    private static readonly (string Bg, string Border, string HeaderFg, string InputBg, string InputFg)[] DropdownRowPalette =
    [
        ("#1E3A8A", "#172554", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#065F46", "#022C22", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#7C2D12", "#431407", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#5B21B6", "#3B0764", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#0E7490", "#164E63", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#854D0E", "#422006", "#FFFFFF", "#F8FAFC", "#0F172A"),
    ];

    public readonly ExcelWorkbookService WorkbookService = new();
    public readonly HeaderProfileStore ProfileStore = new();
    public readonly DataLinkProfileStore DataLinkProfileStore = new();
    public readonly ExcelDataLinkService DataLinkService = new();
    public readonly SessionSettingsStore SessionSettingsStore = new();
    public readonly MultiSheetConfigStore MultiSheetConfigStore = new();
    public readonly AppConfigBackupStore AppConfigBackupStore = new();
    public readonly HsskExportProfileStore HsskExportProfileStore = new();
    public readonly ImportFileExportProfileStore ImportFileExportProfileStore = new();

    public bool IsWorkbookLoaded { get; private set; }
    public string FilePath { get; private set; } = string.Empty;
    public string SelectedSheet { get; private set; } = string.Empty;
    public int HeaderRowNumber { get; set; } = 3;
    public int NameColumnIndex { get; set; } = 2;
    public int SampleRowThreshold { get; set; } = 4;
    public bool AutoSkipBlankHeaders { get; set; } = true;
    public bool IsAutoSaveEnabled { get; set; } = false;
    public bool SuggestionsDisabled { get; set; }
    public string StatusMessage { get; private set; } = "Chọn file Excel để bắt đầu.";
    public bool IsMultiSheetMode { get; private set; }
    public MultiSheetImportSession? MultiSheetSession { get; private set; }
    public string? ProfileKey { get; private set; }
    public List<HeaderDefinition> HeaderColumns { get; } = [];
    public List<RecordItem> Records { get; } = [];
    public List<EditableField> EditableFields { get; } = [];
    public RecordItem? SelectedRecord { get; private set; }
    private readonly Dictionary<CellKey, string> _pendingCellValues = [];
    private int _pendingRowIndex = -1;

    public void OpenWorkbook(string path)
    {
        WorkbookService.OpenWorkbook(path);
        IsWorkbookLoaded = true;
        FilePath = path;
        IsMultiSheetMode = false;
        MultiSheetSession = null;
        ProfileKey = null;
        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
        _pendingCellValues.Clear();
        _pendingRowIndex = -1;
        StatusMessage = $"Đã mở file: {Path.GetFileName(path)}.";

        // Restore saved session settings or defaults (HeaderRow = 3, NameCol = 2, SampleRow = 4)
        var lastSession = SessionSettingsStore.Load();
        HeaderRowNumber = lastSession.HeaderRowNumber > 0 ? lastSession.HeaderRowNumber : 3;
        NameColumnIndex = lastSession.NameColumnIndex > 0 ? lastSession.NameColumnIndex : 2;
        SampleRowThreshold = lastSession.SampleRowThreshold > 0 ? lastSession.SampleRowThreshold : 4;
        AutoSkipBlankHeaders = lastSession.AutoSkipBlankHeaders;
        IsAutoSaveEnabled = false; // Tắt tính năng tự động lưu trên bản web
        SuggestionsDisabled = lastSession.SuggestionsDisabled;

        // Auto-select first sheet just like Windows app with fixed defaults
        var sheets = WorkbookService.GetWorksheetNames();
        if (sheets.Count > 0)
        {
            var firstSheet = sheets[0];
            SelectSheet(firstSheet, HeaderRowNumber, NameColumnIndex, SampleRowThreshold, AutoSkipBlankHeaders, SuggestionsDisabled, false);
        }
    }

    public void CloseWorkbook()
    {
        WorkbookService.DisposePackageOnly();
        IsWorkbookLoaded = false;
        FilePath = string.Empty;
        SelectedSheet = string.Empty;
        IsMultiSheetMode = false;
        MultiSheetSession = null;
        ProfileKey = null;
        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
        _pendingCellValues.Clear();
        _pendingRowIndex = -1;
        StatusMessage = "Chọn file Excel để bắt đầu.";
    }

    public IReadOnlyList<string> GetSheets() => WorkbookService.GetWorksheetNames();
    public IReadOnlyList<WorksheetInfo> GetWorksheetInfos(bool includeHidden) => WorkbookService.GetWorksheetInfos(includeHidden);

    public void SelectSheet(string sheetName, int headerRow, int nameCol, int sampleRow, bool autoSkipBlank = true, bool suggestionsDisabled = false, bool autoSave = true)
    {
        SelectedSheet = sheetName;
        HeaderRowNumber = headerRow > 0 ? headerRow : 3;
        NameColumnIndex = nameCol > 0 ? nameCol : 2;
        SampleRowThreshold = sampleRow > 0 ? sampleRow : 4;
        AutoSkipBlankHeaders = autoSkipBlank;
        SuggestionsDisabled = suggestionsDisabled;
        IsAutoSaveEnabled = false; // Tắt tự động lưu trên bản web
        IsMultiSheetMode = false;
        MultiSheetSession = null;

        WorkbookService.SelectWorksheet(sheetName);
        WorkbookService.InvalidateActiveSheetCaches();
        ApplyHeaderRow();
        LoadRecords();
        SaveSession();
        StatusMessage = $"Đã chọn sheet {sheetName}.";
    }

    private void ApplyHeaderRow()
    {
        var headers = WorkbookService.ReadHeaderRow(HeaderRowNumber);
        if (AutoSkipBlankHeaders)
            headers = headers.Where(h => !string.Equals(h.Name.Trim(), "blank", StringComparison.OrdinalIgnoreCase)).ToList();

        HeaderColumns.Clear();
        foreach (var item in headers)
            HeaderColumns.Add(new HeaderDefinition { ColumnIndex = item.ColumnIndex, Name = item.Name, IsVisible = true });

        ProfileKey = ProfileStore.BuildKey(WorkbookService.WorkbookPath, SelectedSheet, HeaderColumns.Select(h => h.Name));
        ApplyProfileIfAny();
    }

    private void ApplyProfileIfAny()
    {
        if (string.IsNullOrWhiteSpace(ProfileKey)) return;
        var hiddenSet = ProfileStore.TryLoadHiddenHeaders(ProfileKey);
        if (hiddenSet is null) return;
        foreach (var h in HeaderColumns)
            h.IsVisible = !hiddenSet.Contains(h.Name);
    }

    public void LoadRecords()
    {
        if (!IsWorkbookLoaded || HeaderColumns.Count == 0 || NameColumnIndex <= 0) return;
        if (IsMultiSheetMode) { LoadRecordsMultiSheet(); return; }

        var rows = WorkbookService.GetRecordRows(HeaderRowNumber, NameColumnIndex,
            HeaderColumns.Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
        Records.Clear();
        foreach (var row in rows)
            Records.Add(new RecordItem { RowIndex = row.RowIndex, KeyDisplay = row.KeyDisplay, TooltipPreview = row.TooltipPreview });
        SelectedRecord = Records.FirstOrDefault();
        if (SelectedRecord is not null)
        {
            LoadEditableFieldsForSelectedRecord();
        }
        StatusMessage = $"Đã tải {Records.Count} dòng.";
    }

    private void LoadRecordsMultiSheet(bool preserveSelection = false)
    {
        if (MultiSheetSession is null) return;
        var nameSheet = MultiSheetSession.NameSheet;
        if (nameSheet is null) return;

        var nameColumnIndex = MultiSheetSession.NameColumnIndex;
        var nameHeaders = HeaderColumns
            .Where(h => string.Equals(h.SheetName, nameSheet.SheetName, StringComparison.Ordinal))
            .Select(h => new HeaderCell(h.ColumnIndex, h.Name))
            .ToList();
        var prevLogical = preserveSelection ? SelectedRecord?.LogicalIndex : null;
        var rows = WorkbookService.GetLogicalRecordRowsFromSheet(nameSheet.SheetName, nameSheet.FirstDataRow, nameColumnIndex, nameHeaders);
        Records.Clear();
        foreach (var row in rows)
            Records.Add(new RecordItem { LogicalIndex = row.LogicalIndex, RowIndex = row.RowIndex, KeyDisplay = row.KeyDisplay, TooltipPreview = row.TooltipPreview });
        SelectedRecord = prevLogical is int idx ? Records.FirstOrDefault(r => r.LogicalIndex == idx) ?? Records.FirstOrDefault() : Records.FirstOrDefault();
    }

    public void ApplyMultiSheetImport(MultiSheetImportSession session)
    {
        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
        MultiSheetSession = session;
        SelectedSheet = session.NameSheetName;
        WorkbookService.SelectWorksheet(session.NameSheetName);
        IsMultiSheetMode = true;
        BuildMergedHeaderColumns();
        LoadRecordsMultiSheet();
        MultiSheetConfigStore.Save(WorkbookService.WorkbookPath, BuildSavedMultiSheetConfig(session));
        StatusMessage = $"Đang nhập {session.Sheets.Count} sheet ghép.";
    }

    private void BuildMergedHeaderColumns()
    {
        if (MultiSheetSession is null) return;
        HeaderColumns.Clear();
        foreach (var sheetCfg in MultiSheetSession.Sheets)
        {
            var headers = WorkbookService.ReadHeaderBlock(sheetCfg.SheetName, sheetCfg.HeaderFirstRow, sheetCfg.HeaderLastRow);
            foreach (var hc in headers)
            {
                if (MultiSheetSession.AutoSkipBlankHeaders && string.Equals(hc.Name.Trim(), "blank", StringComparison.OrdinalIgnoreCase)) continue;
                HeaderColumns.Add(new HeaderDefinition
                {
                    SheetName = sheetCfg.SheetName,
                    ColumnIndex = hc.ColumnIndex,
                    Name = hc.Name,
                    HeaderFirstRow = sheetCfg.HeaderFirstRow,
                    HeaderLastRow = sheetCfg.HeaderLastRow,
                    IsVisible = true
                });
            }
        }
        ProfileKey = ProfileStore.BuildKey(WorkbookService.WorkbookPath,
            "multi:" + string.Join("|", MultiSheetSession.Sheets.Select(s => s.SheetName)),
            HeaderColumns.Select(h => $"{h.SheetName}:{h.Name}"));
        ApplyProfileIfAny();
    }

    public void ExitMultiSheetMode()
    {
        if (!IsMultiSheetMode) return;
        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        MultiSheetSession = null;
        IsMultiSheetMode = false;
        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
        if (IsWorkbookLoaded && !string.IsNullOrWhiteSpace(SelectedSheet))
        {
            WorkbookService.SelectWorksheet(SelectedSheet);
            ApplyHeaderRow();
            LoadRecords();
        }
    }

    public SavedMultiSheetConfig? TryGetSavedMultiSheetConfig()
    {
        if (!IsWorkbookLoaded || string.IsNullOrWhiteSpace(WorkbookService.WorkbookPath)) return null;
        return MultiSheetConfigStore.TryLoad(WorkbookService.WorkbookPath);
    }

    private static SavedMultiSheetConfig BuildSavedMultiSheetConfig(MultiSheetImportSession session) => new()
    {
        NameSheetName = session.NameSheetName,
        NameColumnIndex = session.NameColumnIndex,
        ShowHiddenSheets = session.ShowHiddenSheets,
        AutoSkipBlankHeaders = session.AutoSkipBlankHeaders,
        SuggestionsDisabled = session.SuggestionsDisabled,
        Sheets = session.Sheets.Select(s => new SavedSheetImportConfig
        {
            SheetName = s.SheetName,
            HeaderFirstRow = s.HeaderFirstRow,
            HeaderLastRow = s.HeaderLastRow,
            FirstDataRow = s.FirstDataRow,
            SampleRow = s.SampleRow,
            IsNameSheet = string.Equals(s.SheetName, session.NameSheetName, StringComparison.Ordinal),
            NameColumnIndex = string.Equals(s.SheetName, session.NameSheetName, StringComparison.Ordinal)
                ? session.NameColumnIndex : s.NameColumnIndex
        }).ToList()
    };

    public void SelectRecord(int rowIndex, int logicalIndex = -1)
    {
        var newRecord = logicalIndex >= 0
            ? Records.FirstOrDefault(r => r.LogicalIndex == logicalIndex)
            : Records.FirstOrDefault(r => r.RowIndex == rowIndex);
        if (newRecord is null) return;

        if (!IsSameRecord(SelectedRecord, newRecord))
        {
            FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
            _pendingRowIndex = -1;
        }
        SelectedRecord = newRecord;
        LoadEditableFieldsForSelectedRecord();
    }

    public void LoadEditableFieldsForSelectedRecord()
    {
        EditableFields.Clear();
        if (SelectedRecord is null) return;
        var visibleHeaders = HeaderColumns.Where(h => h.IsVisible).ToList();

        var singleSheetValues = IsMultiSheetMode
            ? null
            : WorkbookService.ReadRowValues(SelectedRecord.RowIndex, visibleHeaders.Select(h => h.ColumnIndex));

        var fieldDrafts = new List<(HeaderDefinition Header, int SourceRow, string SheetName, string CellValue,
            IReadOnlyList<string> DropdownOptions, IReadOnlyList<int> ParentDropdownColumns, IReadOnlyList<string> SuggestionOptions)>();

        var rowValueOverrides = new Dictionary<CellKey, string>();
        var dropdownColorSlot = 0;
        var dropdownColorGroups = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // First pass: read cell values
        foreach (var header in visibleHeaders)
        {
            var sourceRow = IsMultiSheetMode && MultiSheetSession is not null
                ? MultiSheetSession.GetDataRowForLogicalIndex(header.SheetName, SelectedRecord.LogicalIndex)
                : SelectedRecord.RowIndex;
            var sheetName = IsMultiSheetMode ? header.SheetName : (WorkbookService.ActiveSheetName ?? string.Empty);
            var values = IsMultiSheetMode
                ? WorkbookService.ReadRowValues(sheetName, sourceRow, [header.ColumnIndex])
                : singleSheetValues!;
            var cellValue = values.TryGetValue(header.ColumnIndex, out var v) ? v : string.Empty;
            rowValueOverrides[new CellKey(sheetName, sourceRow, header.ColumnIndex)] = cellValue;
        }

        // Second pass: build fields with dropdown info
        foreach (var header in visibleHeaders)
        {
            var sourceRow = IsMultiSheetMode && MultiSheetSession is not null
                ? MultiSheetSession.GetDataRowForLogicalIndex(header.SheetName, SelectedRecord.LogicalIndex)
                : SelectedRecord.RowIndex;
            var sheetName = IsMultiSheetMode ? header.SheetName : (WorkbookService.ActiveSheetName ?? string.Empty);
            var dependency = WorkbookService.GetDropdownDependencyInfo(sheetName, sourceRow, header.ColumnIndex);
            var dropdownOptions = dependency.IsDependent
                ? WorkbookService.GetDropdownOptions(sheetName, sourceRow, header.ColumnIndex, rowValueOverrides)
                : IsMultiSheetMode
                    ? WorkbookService.GetDropdownOptions(sheetName, sourceRow, header.ColumnIndex)
                    : WorkbookService.GetDropdownOptions(SelectedRecord.RowIndex, header.ColumnIndex);
            dropdownOptions = NormalizeDropdownOptions(dropdownOptions);

            if (dropdownOptions.Count > 0)
            {
                var key = BuildDropdownColorGroupKey(sheetName, header.Name);
                if (!dropdownColorGroups.ContainsKey(key))
                    dropdownColorGroups[key] = dropdownColorSlot++;
            }

            var sampleRow = GetSampleRowForSheet(header.SheetName);
            var suggestionOptions = !SuggestionsDisabled
                && dropdownOptions.Count == 0
                && !dependency.IsDependent
                && (IsMultiSheetMode
                    ? WorkbookService.ColumnHasSuggestionData(header.SheetName, header.ColumnIndex, sampleRow)
                    : WorkbookService.ColumnHasSuggestionData(header.ColumnIndex, sampleRow))
                ? IsMultiSheetMode
                    ? WorkbookService.GetColumnSuggestions(header.SheetName, header.ColumnIndex, sampleRow)
                    : WorkbookService.GetColumnSuggestions(header.ColumnIndex, sampleRow)
                : (IReadOnlyList<string>)[];

            var cellValue = rowValueOverrides[new CellKey(sheetName, sourceRow, header.ColumnIndex)];
            fieldDrafts.Add((header, sourceRow, sheetName, cellValue, dropdownOptions, dependency.ParentColumns, suggestionOptions));
        }

        // Build grouping info
        var headerGroupIndices = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < fieldDrafts.Count; i++)
        {
            var (header, _, sheetName, _, _, _, _) = fieldDrafts[i];
            if (!header.Name.Contains('\n', StringComparison.Ordinal)) continue;
            var key = BuildDropdownColorGroupKey(sheetName, header.Name);
            if (!headerGroupIndices.TryGetValue(key, out var indices)) { indices = []; headerGroupIndices[key] = indices; }
            indices.Add(i);
        }

        var shownParentHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string? lastSheetName = null;
        for (var fieldIndex = 0; fieldIndex < fieldDrafts.Count; fieldIndex++)
        {
            var (header, sourceRow, sheetName, cellValue, dropdownOptions, parentDropdownColumns, suggestionOptions) = fieldDrafts[fieldIndex];
            var rowBg = "Transparent"; var rowBorder = "Transparent"; var headerFg = "#111827";
            var inputBg = "#FFFFFF"; var inputFg = "#111827";
            var groupBorderHex = "Transparent"; var parentTitleBg = "Transparent"; var parentTitleFg = "#111827";
            var colorGroupKey = BuildDropdownColorGroupKey(sheetName, header.Name);
            var hasDropdownColor = dropdownColorGroups.TryGetValue(colorGroupKey, out var paletteSlot);
            var isGroupedHeader = header.Name.Contains('\n', StringComparison.Ordinal);

            if (!isGroupedHeader && hasDropdownColor)
            {
                var p = DropdownRowPalette[paletteSlot % DropdownRowPalette.Length];
                rowBg = p.Bg; rowBorder = p.Border; headerFg = p.HeaderFg; inputBg = p.InputBg; inputFg = p.InputFg;
            }

            string headerDisplayName; string parentHeaderName = string.Empty;
            var showParentHeader = false; var isFirstInHeaderGroup = false; var isLastInHeaderGroup = false;
            if (isGroupedHeader)
            {
                parentHeaderName = ExtractParentHeaderKey(header.Name);
                headerDisplayName = ExtractChildHeaderName(header.Name);
                if (string.IsNullOrWhiteSpace(headerDisplayName)) headerDisplayName = header.Name;
                showParentHeader = shownParentHeaders.Add(colorGroupKey);
                if (headerGroupIndices.TryGetValue(colorGroupKey, out var indices) && indices.Count > 0)
                {
                    isFirstInHeaderGroup = indices[0] == fieldIndex;
                    isLastInHeaderGroup = indices[^1] == fieldIndex;
                }
                groupBorderHex = hasDropdownColor ? DropdownRowPalette[paletteSlot % DropdownRowPalette.Length].Border : "#D1D5DB";
                if (showParentHeader && hasDropdownColor)
                {
                    var palette = DropdownRowPalette[paletteSlot % DropdownRowPalette.Length];
                    parentTitleBg = palette.Bg; parentTitleFg = palette.HeaderFg;
                }
            }
            else { headerDisplayName = header.Name; }

            var showSheetSeparator = IsMultiSheetMode && !string.IsNullOrEmpty(sheetName)
                && !string.Equals(lastSheetName, sheetName, StringComparison.Ordinal);
            if (!string.IsNullOrEmpty(sheetName)) lastSheetName = sheetName;

            var field = new EditableField
            {
                HeaderName = header.Name, HeaderDisplayName = headerDisplayName,
                ParentHeaderName = parentHeaderName, ShowParentHeader = showParentHeader,
                IsGroupedUnderParentHeader = isGroupedHeader, IsFirstInHeaderGroup = isFirstInHeaderGroup,
                IsLastInHeaderGroup = isLastInHeaderGroup, GroupBorderHex = groupBorderHex,
                ParentTitleBackgroundHex = parentTitleBg, ParentTitleForegroundHex = parentTitleFg,
                ColumnIndex = header.ColumnIndex, SheetName = sheetName,
                ShowSheetSeparator = showSheetSeparator, SheetSeparatorTitle = sheetName,
                SourceRowIndex = sourceRow, ParentDropdownColumns = parentDropdownColumns,
                SuggestionOptions = suggestionOptions, Value = cellValue,
                RowHighlightBackgroundHex = rowBg, RowHighlightBorderHex = rowBorder,
                RowHeaderForegroundHex = headerFg, RowInputBackgroundHex = inputBg, RowInputForegroundHex = inputFg
            };
            field.ConfigureDropdownRole(parentDropdownColumns.Count > 0, dropdownOptions);
            field.SetInitialDropdownOptions(dropdownOptions);
            var dropdownSearchIndex = dropdownOptions.Count > DropdownLimits.SearchableThreshold
                ? WorkbookService.GetOrCreateDropdownSearchIndex(sheetName, sourceRow, header.ColumnIndex, rowValueOverrides)
                : null;
            field.InitializeDropdownPresentation(dropdownSearchIndex);
            EditableFields.Add(field);
        }
    }

    public void UpdateCell(string sheetName, int rowIndex, int columnIndex, string value)
    {
        var effectiveSheet = !string.IsNullOrEmpty(sheetName) && WorkbookService.GetWorksheetNames().Contains(sheetName)
            ? sheetName
            : (WorkbookService.ActiveSheetName ?? string.Empty);
        var cellKey = new CellKey(effectiveSheet, rowIndex, columnIndex);
        if (_pendingRowIndex == -1) _pendingRowIndex = rowIndex;
        _pendingCellValues[cellKey] = value;
        // Update in-memory field value
        var field = EditableFields.FirstOrDefault(f =>
            f.ColumnIndex == columnIndex && f.SourceRowIndex == rowIndex &&
            string.Equals(f.SheetName, effectiveSheet, StringComparison.Ordinal));
        field?.SetValueSilently(value);

        var activeNameCol = IsMultiSheetMode && MultiSheetSession is not null
            ? MultiSheetSession.NameColumnIndex : NameColumnIndex;
        if (columnIndex == activeNameCol)
        {
            var targetRec = Records.FirstOrDefault(r => r.RowIndex == rowIndex);
            if (targetRec != null)
            {
                targetRec.KeyDisplay = string.IsNullOrWhiteSpace(value) ? $"(Dòng {rowIndex})" : value.Trim();
            }
            if (SelectedRecord != null && SelectedRecord.RowIndex == rowIndex)
            {
                SelectedRecord.KeyDisplay = string.IsNullOrWhiteSpace(value) ? $"(Dòng {rowIndex})" : value.Trim();
            }
        }

        if (IsAutoSaveEnabled) FlushPendingChanges(saveToDisk: true);
    }

    public void FlushAndSave()
    {
        FlushPendingChanges(saveToDisk: true);
    }

    private void FlushPendingChanges(bool saveToDisk)
    {
        if (_pendingRowIndex == -1 || _pendingCellValues.Count == 0)
        {
            if (saveToDisk && WorkbookService.HasUnsavedChanges())
            {
                if (WorkbookService.TrySave(out var e)) StatusMessage = $"Đã lưu file lúc {DateTime.Now:HH:mm:ss}.";
                else StatusMessage = e ?? "Lưu thất bại.";
            }
            return;
        }

        try
        {
            foreach (var cell in _pendingCellValues)
            {
                var sheet = !string.IsNullOrEmpty(cell.Key.SheetName) && WorkbookService.GetWorksheetNames().Contains(cell.Key.SheetName)
                    ? cell.Key.SheetName
                    : WorkbookService.ActiveSheetName;
                if (string.IsNullOrEmpty(sheet)) continue;

                var oldValue = WorkbookService.ReadCellText(sheet, cell.Key.RowIndex, cell.Key.ColumnIndex);
                WorkbookService.UpdateCellValue(sheet, cell.Key.RowIndex, cell.Key.ColumnIndex, cell.Value);
                var sampleRow = GetSampleRowForSheet(sheet);
                WorkbookService.UpdateSuggestionCacheAfterCellEdit(sheet, cell.Key.RowIndex, cell.Key.ColumnIndex, sampleRow, oldValue, cell.Value);
            }
            if (saveToDisk)
            {
                if (WorkbookService.TrySave(out var e)) StatusMessage = $"Đã lưu dòng {_pendingRowIndex} lúc {DateTime.Now:HH:mm:ss}.";
                else StatusMessage = e ?? "Lưu thất bại.";
            }
            else { WorkbookService.CommitPendingEditsToWorksheetOnly(); }
        }
        finally
        {
            _pendingCellValues.Clear();
            _pendingRowIndex = -1;
        }
    }

    public int AddRecord(string nameValue)
    {
        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        int rowIndex;
        if (IsMultiSheetMode && MultiSheetSession is not null)
        {
            var nameSheet = MultiSheetSession.NameSheet!;
            rowIndex = WorkbookService.AddRecordAtSheetLogicalEnd(nameSheet.SheetName, nameSheet.FirstDataRow,
                MultiSheetSession.NameColumnIndex, nameValue,
                HeaderColumns.Where(h => string.Equals(h.SheetName, nameSheet.SheetName, StringComparison.Ordinal))
                    .Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
            LoadRecordsMultiSheet(preserveSelection: false);
        }
        else
        {
            rowIndex = WorkbookService.AddRecordAtLogicalEnd(HeaderRowNumber, NameColumnIndex, nameValue,
                HeaderColumns.Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
            LoadRecords();
        }
        SelectRecord(rowIndex);
        if (IsAutoSaveEnabled) WorkbookService.TrySave(out _);
        return rowIndex;
    }

    public void DeleteRecord(int rowIndex, int logicalIndex = -1)
    {
        var record = logicalIndex >= 0
            ? Records.FirstOrDefault(r => r.LogicalIndex == logicalIndex)
            : Records.FirstOrDefault(r => r.RowIndex == rowIndex);
        if (record is null) return;
        _pendingCellValues.Clear();
        _pendingRowIndex = -1;

        if (IsMultiSheetMode && MultiSheetSession is not null)
        {
            foreach (var sheet in MultiSheetSession.Sheets)
            {
                var sheetRow = MultiSheetSession.GetDataRowForLogicalIndex(sheet.SheetName, record.LogicalIndex);
                if (sheetRow > 0) WorkbookService.DeleteRowOnSheet(sheet.SheetName, sheetRow);
            }
            LoadRecordsMultiSheet();
        }
        else
        {
            WorkbookService.DeleteRow(record.RowIndex);
            LoadRecords();
        }
        if (IsAutoSaveEnabled) WorkbookService.TrySave(out _);
    }

    public void RenameRecord(int rowIndex, int logicalIndex, string newName)
    {
        var nameColIdx = IsMultiSheetMode && MultiSheetSession is not null
            ? MultiSheetSession.NameColumnIndex : NameColumnIndex;
        var sheetName = IsMultiSheetMode && MultiSheetSession?.NameSheetName is { } ns ? ns : WorkbookService.ActiveSheetName ?? string.Empty;
        var actualRow = logicalIndex >= 0 && MultiSheetSession?.NameSheet is { } nsh
            ? MultiSheetSession.GetDataRowForLogicalIndex(nsh.SheetName, logicalIndex)
            : rowIndex;
        WorkbookService.UpdateCellValue(sheetName, actualRow, nameColIdx, newName);
        if (IsAutoSaveEnabled) WorkbookService.TrySave(out _);
        if (IsMultiSheetMode) LoadRecordsMultiSheet(preserveSelection: true);
        else LoadRecords();
    }

    public bool TrySave(out string? error) => WorkbookService.TrySave(out error);

    public bool TrySaveAs(string newPath)
    {
        FlushPendingChanges(saveToDisk: false);
        return WorkbookService.TrySaveAs(newPath, out _);
    }

    public IReadOnlyList<string> GetDropdownOptions(int columnIndex, int rowIndex, string sheetName)
        => NormalizeDropdownOptions(WorkbookService.GetDropdownOptions(rowIndex, columnIndex));

    public IReadOnlyList<string> GetDependentDropdownOptions(string sheetName, int rowIndex, int columnIndex,
        Dictionary<CellKey, string> overrides)
        => NormalizeDropdownOptions(WorkbookService.GetDropdownOptions(sheetName, rowIndex, columnIndex, overrides));

    public IReadOnlyList<string> SearchDropdownOptions(string sheetName, int rowIndex, int columnIndex,
        string filter, Dictionary<CellKey, string>? overrides = null)
    {
        var emptyOverrides = new Dictionary<CellKey, string>();
        var index = overrides != null
            ? WorkbookService.GetOrCreateDropdownSearchIndex(sheetName, rowIndex, columnIndex, overrides)
            : WorkbookService.GetOrCreateDropdownSearchIndex(sheetName, rowIndex, columnIndex, emptyOverrides);
        return index.Search(filter, DropdownLimits.UnlimitedVisibleItems, null);
    }

    public void SetHeaderVisibility(int columnIndex, string sheetName, bool isVisible)
    {
        var header = HeaderColumns.FirstOrDefault(h =>
            h.ColumnIndex == columnIndex && string.Equals(h.SheetName, sheetName, StringComparison.Ordinal));
        if (header is null) return;
        header.IsVisible = isVisible;
        if (!string.IsNullOrWhiteSpace(ProfileKey))
        {
            var hiddenHeaders = HeaderColumns.Where(h => !h.IsVisible).Select(h => h.Name).ToList();
            ProfileStore.SaveHiddenHeaders(ProfileKey, hiddenHeaders);
        }
    }

    public void ShowAllHeaders(bool show)
    {
        foreach (var h in HeaderColumns) h.IsVisible = show;
        if (!string.IsNullOrWhiteSpace(ProfileKey))
        {
            var hiddenHeaders = show ? new List<string>() : HeaderColumns.Select(h => h.Name).ToList();
            ProfileStore.SaveHiddenHeaders(ProfileKey, hiddenHeaders);
        }
    }

    private bool _isShowAllHeadersTemporarilyEnabled;
    private HashSet<string> _hiddenHeadersBeforeShowAll = [];

    public string ToggleAllHeaders()
    {
        if (_isShowAllHeadersTemporarilyEnabled)
        {
            foreach (var h in HeaderColumns)
                h.IsVisible = !_hiddenHeadersBeforeShowAll.Contains(h.Name);
            _isShowAllHeadersTemporarilyEnabled = false;
        }
        else
        {
            _hiddenHeadersBeforeShowAll = HeaderColumns.Where(h => !h.IsVisible).Select(h => h.Name).ToHashSet();
            foreach (var h in HeaderColumns)
                h.IsVisible = true;
            _isShowAllHeadersTemporarilyEnabled = true;
        }
        LoadEditableFieldsForSelectedRecord();
        return _isShowAllHeadersTemporarilyEnabled ? "Tắt hiện tất cả" : "Hiện tất cả";
    }

    public string GetToggleAllHeadersLabel() => _isShowAllHeadersTemporarilyEnabled ? "Tắt hiện tất cả" : "Hiện tất cả";

    public HsskExportProfile GetHsskExportProfile()
    {
        if (string.IsNullOrWhiteSpace(ProfileKey)) return new HsskExportProfile();
        return HsskExportProfileStore.TryLoad(ProfileKey) ?? new HsskExportProfile();
    }

    public void SaveHsskExportProfile(HsskExportProfile profile)
    {
        if (string.IsNullOrWhiteSpace(ProfileKey)) return;
        HsskExportProfileStore.Save(ProfileKey, profile);
    }

    public ImportFileExportProfile GetImportFileExportProfile()
    {
        if (string.IsNullOrWhiteSpace(ProfileKey)) return new ImportFileExportProfile();
        return ImportFileExportProfileStore.TryLoad(ProfileKey) ?? new ImportFileExportProfile();
    }

    public void SaveImportFileExportProfile(ImportFileExportProfile profile)
    {
        if (string.IsNullOrWhiteSpace(ProfileKey)) return;
        ImportFileExportProfileStore.Save(ProfileKey, profile);
    }

    public void SortRecords(int columnIndex, ColumnSortMode mode, IReadOnlyList<string>? customValues = null)
    {
        if (!IsWorkbookLoaded) return;
        var startRow = HeaderRowNumber + 1;
        var excludedRow = SampleRowThreshold > 0 ? SampleRowThreshold : 0;
        WorkbookService.SortRowsByColumn(startRow, columnIndex, mode, customValues, excludedRow);
        WorkbookService.TrySave(out _);
        LoadRecords();
    }

    private void SaveSession()
    {
        SessionSettingsStore.Save(new SessionSettings
        {
            HeaderRowNumber = HeaderRowNumber,
            NameColumnIndex = NameColumnIndex,
            SampleRowThreshold = SampleRowThreshold,
            AutoSkipBlankHeaders = AutoSkipBlankHeaders,
            IsAutoSaveEnabled = IsAutoSaveEnabled,
            SuggestionsDisabled = SuggestionsDisabled
        });
    }

    private int GetSampleRowForSheet(string sheetName)
    {
        if (IsMultiSheetMode && MultiSheetSession?.GetSheetConfig(sheetName) is { } cfg)
            return ResolveSampleRow(cfg.SampleRow, cfg.FirstDataRow, cfg.HeaderLastRow);
        return SampleRowThreshold;
    }

    private static int ResolveSampleRow(int configuredSampleRow, int firstDataRow, int headerLastRow)
    {
        if (configuredSampleRow > 0) return configuredSampleRow;
        return firstDataRow > headerLastRow ? firstDataRow : Math.Max(1, headerLastRow);
    }

    private static bool IsSameRecord(RecordItem? left, RecordItem? right)
    {
        if (left is null && right is null) return true;
        if (left is null || right is null) return false;
        if (left.LogicalIndex >= 0 && right.LogicalIndex >= 0) return left.LogicalIndex == right.LogicalIndex;
        return left.RowIndex == right.RowIndex;
    }

    private static IReadOnlyList<string> NormalizeDropdownOptions(IReadOnlyList<string> opts)
    {
        if (opts.Count == 0) return opts;
        var n = opts.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct(StringComparer.Ordinal).ToList();
        n.Add(string.Empty);
        return n;
    }

    private static string BuildDropdownColorGroupKey(string sheetName, string headerName)
    {
        var parentKey = headerName.Contains('\n') ? ExtractParentHeaderKey(headerName) : headerName;
        return string.IsNullOrEmpty(sheetName) ? parentKey : $"{sheetName}:{parentKey}";
    }

    private static string ExtractParentHeaderKey(string name)
    {
        var idx = name.IndexOf('\n');
        return idx >= 0 ? name[..idx] : name;
    }

    private static string ExtractChildHeaderName(string name)
    {
        var idx = name.LastIndexOf('\n');
        return idx >= 0 ? name[(idx + 1)..] : name;
    }

    public string? RunDataLink(DataLinkProfile profile)
    {
        if (!IsWorkbookLoaded) return "Chưa mở file Excel.";
        if (HeaderColumns.Count == 0) return "Chưa có tiêu đề cột.";
        if (string.IsNullOrWhiteSpace(profile.DataFilePath) || !File.Exists(profile.DataFilePath))
            return "File số liệu không tồn tại.";
        if (string.IsNullOrWhiteSpace(profile.DataSheetName)) return "Chưa chọn sheet file số liệu.";
        if (profile.MainKeyColumnIndex <= 0 || profile.DataKeyColumnIndex <= 0) return "Chưa chọn cột key.";
        if (profile.ColumnMappings.Count == 0) return "Chưa có cặp cột ánh xạ nào.";

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);

        DataLinkMainContext mainContext;
        if (IsMultiSheetMode && MultiSheetSession is not null)
        {
            var keySheetName = string.IsNullOrWhiteSpace(profile.MainKeySheetName)
                ? MultiSheetSession.NameSheetName : profile.MainKeySheetName;
            var keySheetCfg = MultiSheetSession.GetSheetConfig(keySheetName);
            if (keySheetCfg is null) return $"Không tìm thấy cấu hình sheet key '{keySheetName}'.";
            mainContext = new DataLinkMainContext
            {
                KeySheetName = keySheetName,
                KeySheetFirstDataRow = keySheetCfg.FirstDataRow,
                FirstDataRowBySheet = MultiSheetSession.Sheets.ToDictionary(s => s.SheetName, s => s.FirstDataRow, StringComparer.Ordinal),
                FormatReferenceRowBySheet = MultiSheetSession.Sheets.ToDictionary(s => s.SheetName, s => GetSampleRowForSheet(s.SheetName), StringComparer.Ordinal)
            };
        }
        else
        {
            var afterHeader = HeaderRowNumber + 1;
            var firstDataRow = SampleRowThreshold > 0 ? Math.Max(afterHeader, SampleRowThreshold) : afterHeader;
            mainContext = new DataLinkMainContext { KeySheetName = string.Empty, KeySheetFirstDataRow = firstDataRow };
        }

        var dataFirstDataRow = Math.Max(profile.DataHeaderRow + 1, 1);
        var formatReferenceRow = SampleRowThreshold > 0 ? SampleRowThreshold : Math.Max(1, HeaderRowNumber);
        try
        {
            using var reader = new ExcelLookupReader();
            reader.Open(profile.DataFilePath, profile.DataSheetName);
            var result = DataLinkService.RunLink(WorkbookService, reader, mainContext, dataFirstDataRow,
                profile.MainKeyColumnIndex, profile.DataKeyColumnIndex, profile.ColumnMappings, formatReferenceRow);
            if (IsAutoSaveEnabled) WorkbookService.TrySave(out _);
            else WorkbookService.CommitPendingEditsToWorksheetOnly();
            LoadRecords();
            StatusMessage = $"Đã liên kết {result.RowsMatched} dòng, ghi {result.CellsWritten} ô.";
            return null;
        }
        catch (Exception ex) { return $"Không thể liên kết số liệu: {ex.Message}"; }
    }

    public void Dispose() => WorkbookService.Dispose();
}
