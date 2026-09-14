using ExcelDataEntryApp.Models;
using OfficeOpenXml;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Table;

namespace ExcelDataEntryApp.Services;

public sealed class ExcelWorkbookService : IDisposable
{
    private readonly ExcelInteropDropdownReader _interopDropdownReader = new();
    private readonly DynamicDropdownResolver _dynamicDropdownResolver = new();
    private readonly ExtendedDataValidationCatalog _extendedValidations = new();
    private readonly Dictionary<(string Sheet, int Column, int Sample), SuggestionCacheEntry> _suggestionCache = [];
    private readonly Dictionary<(string Sheet, int Row, int Column), IReadOnlyList<string>> _dropdownCellCache = [];
    private readonly Dictionary<(string Sheet, int Column), IReadOnlyList<string>> _dropdownColumnCache = [];
    private readonly Dictionary<(string Sheet, int Column), DropdownSearchIndex> _dropdownSearchIndexCache = [];
    private ExcelPackage? _package;
    private ExcelWorksheet? _worksheet;
    private string _workbookPath = string.Empty;
    private bool _hasUnsavedChanges;

    public string WorkbookPath => _workbookPath;
    public string? ActiveSheetName => _worksheet?.Name;

    public void OpenWorkbook(string path)
    {
        DisposePackageOnly();
        ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
        _workbookPath = path;
        _package = new ExcelPackage(new FileInfo(path));
        _worksheet = null;
        _hasUnsavedChanges = false;
        _suggestionCache.Clear();
        _dropdownCellCache.Clear();
        _dropdownColumnCache.Clear();
        _dropdownSearchIndexCache.Clear();
        _extendedValidations.LoadFromWorkbook(path);
    }

    public IReadOnlyList<string> GetWorksheetNames()
    {
        return GetWorksheetInfos(includeHidden: true).Select(w => w.Name).ToList();
    }

    public IReadOnlyList<WorksheetInfo> GetWorksheetInfos(bool includeHidden)
    {
        if (_package is null)
        {
            return [];
        }

        return _package.Workbook.Worksheets
            .Where(ws => includeHidden || ws.Hidden == eWorkSheetHidden.Visible)
            .Select(ws => new WorksheetInfo(ws.Name, ws.Hidden != eWorkSheetHidden.Visible))
            .ToList();
    }

    public void SelectWorksheet(string sheetName)
    {
        _worksheet = _package?.Workbook.Worksheets[sheetName];
        if (_worksheet is null)
        {
            throw new InvalidOperationException("Không tìm thấy sheet đã chọn.");
        }
    }

    public IReadOnlyList<int> GetCandidateHeaderRows(int maxRows = 200)
    {
        var endRow = _worksheet?.Dimension?.End.Row ?? 0;
        return Enumerable.Range(1, Math.Max(0, Math.Min(maxRows, endRow))).ToList();
    }

    public IReadOnlyList<HeaderCell> ReadHeaderRow(int headerRow)
    {
        EnsureWorksheet();
        var endCol = _worksheet!.Dimension?.End.Column ?? 0;
        var mergeResolver = new HeaderMergeResolver(_worksheet);
        var headers = new List<HeaderCell>();
        for (var col = 1; col <= endCol; col++)
        {
            var text = mergeResolver.GetEffectiveText(headerRow, col)?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            headers.Add(new HeaderCell(col, text));
        }

        return headers;
    }

    public IReadOnlyList<HeaderCell> ReadHeaderBlock(string sheetName, int firstRow, int lastRow)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var startRow = Math.Max(1, firstRow);
        var endHeaderRow = Math.Max(startRow, lastRow);
        var endCol = worksheet.Dimension?.End.Column ?? 0;
        var mergeResolver = new HeaderMergeResolver(worksheet);
        var headers = new List<HeaderCell>();
        for (var col = 1; col <= endCol; col++)
        {
            var parts = mergeResolver.BuildColumnHeaderParts(col, startRow, endHeaderRow);
            if (parts.Count == 0)
            {
                continue;
            }

            headers.Add(new HeaderCell(col, string.Join('\n', parts)));
        }

        return headers;
    }

    public int GetWorksheetColumnCount(string sheetName)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        return worksheet.Dimension?.End.Column ?? 0;
    }

    public IReadOnlyList<LogicalRecordRow> GetLogicalRecordRowsFromSheet(
        string sheetName,
        int firstDataRow,
        int idColumn,
        IReadOnlyList<HeaderCell> headers)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var endRow = worksheet.Dimension?.End.Row ?? 0;
        var rows = new List<LogicalRecordRow>();
        var previewColumns = headers.Where(h => h.ColumnIndex != idColumn).Take(4).ToList();
        var logicalIndex = 0;

        for (var row = firstDataRow; row <= endRow; row++)
        {
            var keyDisplay = worksheet.Cells[row, idColumn].Text?.Trim() ?? string.Empty;
            var hasAnyValue = headers.Any(h => !string.IsNullOrWhiteSpace(worksheet.Cells[row, h.ColumnIndex].Text));
            if (!hasAnyValue)
            {
                continue;
            }

            var safeKey = string.IsNullOrWhiteSpace(keyDisplay) ? $"(Dòng {row})" : keyDisplay;
            var previewParts = previewColumns
                .Select(h => $"{h.Name.Replace('\n', ' ')}: {worksheet.Cells[row, h.ColumnIndex].Text}")
                .Where(v => !v.EndsWith(": ", StringComparison.Ordinal))
                .ToList();

            rows.Add(new LogicalRecordRow(logicalIndex++, row, safeKey, string.Join(" | ", previewParts)));
        }

        return rows;
    }

    public IReadOnlyList<RecordRow> GetRecordRows(int headerRow, int idColumn, IReadOnlyList<HeaderCell> headers)
    {
        EnsureWorksheet();
        var endRow = _worksheet!.Dimension?.End.Row ?? 0;
        var rows = new List<RecordRow>();
        var previewColumns = headers.Where(h => h.ColumnIndex != idColumn).Take(4).ToList();

        for (var row = headerRow + 1; row <= endRow; row++)
        {
            var keyDisplay = _worksheet.Cells[row, idColumn].Text?.Trim() ?? string.Empty;
            var hasAnyValue = headers.Any(h => !string.IsNullOrWhiteSpace(_worksheet.Cells[row, h.ColumnIndex].Text));
            if (!hasAnyValue)
            {
                continue;
            }

            var safeKey = string.IsNullOrWhiteSpace(keyDisplay) ? $"(Dòng {row})" : keyDisplay;
            var previewParts = previewColumns
                .Select(h => $"{h.Name}: {_worksheet.Cells[row, h.ColumnIndex].Text}")
                .Where(v => !v.EndsWith(": ", StringComparison.Ordinal))
                .ToList();

            rows.Add(new RecordRow(row, safeKey, string.Join(" | ", previewParts)));
        }

        return rows;
    }

    public int AddRecordAtLogicalEnd(int headerRow, int nameColumnIndex, string nameValue, IReadOnlyList<HeaderCell> headers)
    {
        EnsureWorksheet();
        var endRow = _worksheet!.Dimension?.End.Row ?? 0;
        var endColumn = _worksheet.Dimension?.End.Column ?? 0;
        var lastDataRow = headerRow;
        var scanColumnIndexes = headers.Count > 0
            ? headers.Select(h => h.ColumnIndex).Distinct().ToList()
            : new List<int>();
        if (!scanColumnIndexes.Contains(nameColumnIndex))
        {
            scanColumnIndexes.Add(nameColumnIndex);
        }
        if (endColumn > 0)
        {
            for (var col = 1; col <= endColumn; col++)
            {
                if (!scanColumnIndexes.Contains(col))
                {
                    scanColumnIndexes.Add(col);
                }
            }
        }

        for (var row = headerRow + 1; row <= endRow; row++)
        {
            var hasAnyValue = scanColumnIndexes.Any(col => !string.IsNullOrWhiteSpace(_worksheet.Cells[row, col].Text));
            if (hasAnyValue)
            {
                lastDataRow = row;
            }
        }

        var newRow = lastDataRow + 1;
        _worksheet.Cells[newRow, nameColumnIndex].Value = nameValue;
        _hasUnsavedChanges = true;
        InvalidateActiveSheetCaches();
        return newRow;
    }

    /// <summary>Xóa một dòng trong sheet hiện tại (1-based row index).</summary>
    public void DeleteRow(int rowIndex)
    {
        EnsureWorksheet();
        DeleteRowOnSheet(_worksheet!.Name, rowIndex);
    }

    public void DeleteRowOnSheet(string sheetName, int rowIndex)
    {
        if (rowIndex < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        }

        var worksheet = GetWorksheetOrThrow(sheetName);
        worksheet.DeleteRow(rowIndex);
        _hasUnsavedChanges = true;
        InvalidateSheetCaches(sheetName);
    }

    public int AddRecordAtSheetLogicalEnd(
        string sheetName,
        int firstDataRow,
        int nameColumnIndex,
        string nameValue,
        IReadOnlyList<HeaderCell> headers)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var endRow = worksheet.Dimension?.End.Row ?? 0;
        var endColumn = worksheet.Dimension?.End.Column ?? 0;
        var lastDataRow = firstDataRow - 1;
        var scanColumnIndexes = headers.Count > 0
            ? headers.Select(h => h.ColumnIndex).Distinct().ToList()
            : [];
        if (!scanColumnIndexes.Contains(nameColumnIndex))
        {
            scanColumnIndexes.Add(nameColumnIndex);
        }

        if (endColumn > 0)
        {
            for (var col = 1; col <= endColumn; col++)
            {
                if (!scanColumnIndexes.Contains(col))
                {
                    scanColumnIndexes.Add(col);
                }
            }
        }

        for (var row = firstDataRow; row <= endRow; row++)
        {
            var hasAnyValue = scanColumnIndexes.Any(col => !string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text));
            if (hasAnyValue)
            {
                lastDataRow = row;
            }
        }

        var newRow = lastDataRow + 1;
        if (!string.IsNullOrEmpty(nameValue) && nameColumnIndex > 0)
        {
            worksheet.Cells[newRow, nameColumnIndex].Value = nameValue;
        }

        _hasUnsavedChanges = true;
        InvalidateSheetCaches(sheetName);
        return newRow;
    }

    public IReadOnlyList<string> GetDistinctColumnValues(int firstDataRow, int sortColumnIndex, int excludedRowIndex = 0)
    {
        EnsureWorksheet();
        if (firstDataRow < 1 || sortColumnIndex < 1)
        {
            return [];
        }

        var endRow = _worksheet!.Dimension?.End.Row ?? 0;
        if (firstDataRow > endRow)
        {
            return [];
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var distinct = new List<string>();
        for (var row = firstDataRow; row <= endRow; row++)
        {
            if (excludedRowIndex > 0 && row == excludedRowIndex)
            {
                continue;
            }

            var keyText = _worksheet.Cells[row, sortColumnIndex].Text?.Trim() ?? string.Empty;
            if (seen.Add(keyText))
            {
                distinct.Add(keyText);
            }
        }

        return distinct
            .OrderBy(v => string.IsNullOrWhiteSpace(v) ? 1 : 0)
            .ThenBy(v => v, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    public void SortRowsByColumn(
        int firstDataRow,
        int sortColumnIndex,
        ColumnSortMode mode,
        IReadOnlyList<string>? customOrder = null,
        int excludedRowIndex = 0)
    {
        EnsureWorksheet();
        if (firstDataRow < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(firstDataRow));
        }

        if (sortColumnIndex < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(sortColumnIndex));
        }

        var endRow = _worksheet!.Dimension?.End.Row ?? 0;
        if (firstDataRow > endRow)
        {
            return;
        }

        var endCol = _worksheet.Dimension?.End.Column ?? 0;
        if (endCol <= 0)
        {
            return;
        }

        var rows = new List<(string SortKey, object?[] Values)>();
        for (var row = firstDataRow; row <= endRow; row++)
        {
            if (excludedRowIndex > 0 && row == excludedRowIndex)
            {
                continue;
            }

            var values = new object?[endCol];
            for (var col = 1; col <= endCol; col++)
            {
                values[col - 1] = _worksheet.Cells[row, col].Value;
            }

            var keyText = _worksheet.Cells[row, sortColumnIndex].Text?.Trim() ?? string.Empty;
            rows.Add((keyText, values));
        }

        if (rows.Count <= 1)
        {
            return;
        }

        var orderedRows = mode switch
        {
            ColumnSortMode.Descending => rows
                .OrderBy(r => string.IsNullOrWhiteSpace(r.SortKey) ? 1 : 0)
                .ThenByDescending(r => r.SortKey, StringComparer.CurrentCultureIgnoreCase)
                .ToList(),
            ColumnSortMode.Custom when customOrder is { Count: > 0 } => OrderRowsByCustomOrder(rows, customOrder),
            _ => rows
                .OrderBy(r => string.IsNullOrWhiteSpace(r.SortKey) ? 1 : 0)
                .ThenBy(r => r.SortKey, StringComparer.CurrentCultureIgnoreCase)
                .ToList()
        };

        var orderedIndex = 0;
        for (var row = firstDataRow; row <= endRow; row++)
        {
            if (excludedRowIndex > 0 && row == excludedRowIndex)
            {
                continue;
            }

            var rowValues = orderedRows[orderedIndex].Values;
            orderedIndex++;
            for (var col = 1; col <= endCol; col++)
            {
                _worksheet.Cells[row, col].Value = rowValues[col - 1];
            }
        }

        _hasUnsavedChanges = true;
        InvalidateActiveSheetCaches();
    }

    private static List<(string SortKey, object?[] Values)> OrderRowsByCustomOrder(
        List<(string SortKey, object?[] Values)> rows,
        IReadOnlyList<string> customOrder)
    {
        var rankMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < customOrder.Count; i++)
        {
            if (!rankMap.ContainsKey(customOrder[i]))
            {
                rankMap[customOrder[i]] = i;
            }
        }

        return rows
            .Select((row, index) => (row, index))
            .OrderBy(x => rankMap.TryGetValue(x.row.SortKey, out var rank) ? rank : int.MaxValue)
            .ThenBy(x => x.index)
            .Select(x => x.row)
            .ToList();
    }

    public Dictionary<int, string> ReadRowValues(int rowIndex, IEnumerable<int> columns)
    {
        EnsureWorksheet();
        return ReadRowValues(_worksheet!.Name, rowIndex, columns);
    }

    public Dictionary<int, string> ReadRowValues(string sheetName, int rowIndex, IEnumerable<int> columns)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        return columns.ToDictionary(c => c, c => worksheet.Cells[rowIndex, c].Text ?? string.Empty);
    }

    public IReadOnlyList<string> ReadAllCellTextsInRow(int rowIndex)
    {
        EnsureWorksheet();
        var endCol = _worksheet!.Dimension?.End.Column ?? 0;
        var values = new List<string>(Math.Max(endCol, 0));
        for (var col = 1; col <= endCol; col++)
        {
            values.Add(_worksheet.Cells[rowIndex, col].Text ?? string.Empty);
        }

        return values;
    }

    public string ReadCellText(int rowIndex, int columnIndex)
    {
        EnsureWorksheet();
        return ReadCellText(_worksheet!.Name, rowIndex, columnIndex);
    }

    public string ReadCellText(string sheetName, int rowIndex, int columnIndex)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        return worksheet.Cells[rowIndex, columnIndex].Text?.Trim() ?? string.Empty;
    }

    public IReadOnlyList<int> FindColumnsWithMarkerInRows(string sheetName, int row1, int row2, string marker)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var endCol = worksheet.Dimension?.End.Column ?? 0;
        if (endCol <= 0)
        {
            return [];
        }

        var columns = new List<int>();
        for (var col = 1; col <= endCol; col++)
        {
            var text1 = worksheet.Cells[row1, col].Text?.Trim() ?? string.Empty;
            var text2 = worksheet.Cells[row2, col].Text?.Trim() ?? string.Empty;
            if (IsMarkerCell(text1, marker) || IsMarkerCell(text2, marker))
            {
                columns.Add(col);
            }
        }

        return columns;
    }

    private static bool IsMarkerCell(string cellText, string marker) =>
        string.Equals(cellText.Trim(), marker, StringComparison.OrdinalIgnoreCase);

    public int GetEndRow()
    {
        EnsureWorksheet();
        return GetEndRow(_worksheet!.Name);
    }

    public int GetEndRow(string sheetName)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        return worksheet.Dimension?.End.Row ?? 0;
    }

    public int FindLastNameRow(string sheetName, int nameColumnIndex, int firstDataRow)
    {
        if (nameColumnIndex < 1 || firstDataRow < 1)
        {
            return -1;
        }

        var endRow = GetEndRow(sheetName);
        var lastNameRow = -1;
        for (var row = firstDataRow; row <= endRow; row++)
        {
            if (!string.IsNullOrWhiteSpace(ReadCellText(sheetName, row, nameColumnIndex)))
            {
                lastNameRow = row;
            }
        }

        return lastNameRow;
    }

    public IReadOnlyList<string> GetColumnSuggestions(int columnIndex, int sampleRowThreshold)
    {
        EnsureWorksheet();
        return GetColumnSuggestions(_worksheet!.Name, columnIndex, sampleRowThreshold);
    }

    public IReadOnlyList<string> GetColumnSuggestions(string sheetName, int columnIndex, int sampleRowThreshold)
    {
        return EnsureSuggestionCache(sheetName, columnIndex, sampleRowThreshold).Values.ToList();
    }

    public bool ColumnHasSuggestionData(int columnIndex, int sampleRowThreshold)
    {
        EnsureWorksheet();
        return ColumnHasSuggestionData(_worksheet!.Name, columnIndex, sampleRowThreshold);
    }

    public bool ColumnHasSuggestionData(string sheetName, int columnIndex, int sampleRowThreshold)
    {
        return EnsureSuggestionCache(sheetName, columnIndex, sampleRowThreshold).Values.Count > 0;
    }

    public bool ColumnContainsSuggestionValue(int columnIndex, int sampleRowThreshold, string value)
    {
        EnsureWorksheet();
        var normalizedTarget = value.Trim();
        if (columnIndex <= 0 || string.IsNullOrWhiteSpace(normalizedTarget) || IsPureNumericText(normalizedTarget))
        {
            return false;
        }

        var endRow = _worksheet!.Dimension?.End.Row ?? 0;
        if (endRow <= sampleRowThreshold)
        {
            return false;
        }

        for (var row = sampleRowThreshold + 1; row <= endRow; row++)
        {
            var text = _worksheet.Cells[row, columnIndex].Text?.Trim() ?? string.Empty;
            if (string.Equals(text, normalizedTarget, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    public void UpdateSuggestionCacheAfterCellEdit(int rowIndex, int columnIndex, int sampleRowThreshold, string oldValue, string newValue)
    {
        EnsureWorksheet();
        UpdateSuggestionCacheAfterCellEdit(_worksheet!.Name, rowIndex, columnIndex, sampleRowThreshold, oldValue, newValue);
    }

    public void UpdateSuggestionCacheAfterCellEdit(
        string sheetName,
        int rowIndex,
        int columnIndex,
        int sampleRowThreshold,
        string oldValue,
        string newValue)
    {
        if (rowIndex <= sampleRowThreshold || columnIndex <= 0)
        {
            return;
        }

        var entry = EnsureSuggestionCache(sheetName, columnIndex, sampleRowThreshold);
        var normalizedOld = NormalizeSuggestionText(oldValue);
        var normalizedNew = NormalizeSuggestionText(newValue);

        if (!string.IsNullOrWhiteSpace(normalizedOld)
            && !string.Equals(normalizedOld, normalizedNew, StringComparison.OrdinalIgnoreCase)
            && !ColumnContainsSuggestionValueOnSheet(sheetName, columnIndex, sampleRowThreshold, normalizedOld))
        {
            entry.Values.RemoveWhere(v => string.Equals(v, normalizedOld, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.IsNullOrWhiteSpace(normalizedNew))
        {
            entry.Values.Add(normalizedNew);
        }
    }

    private static bool IsPureNumericText(string text)
    {
        var normalized = text.Trim().Replace(" ", string.Empty);
        return double.TryParse(normalized, NumberStyles.Number, CultureInfo.CurrentCulture, out _)
            || double.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out _);
    }

    private SuggestionCacheEntry EnsureSuggestionCache(int columnIndex, int sampleRowThreshold)
    {
        EnsureWorksheet();
        return EnsureSuggestionCache(_worksheet!.Name, columnIndex, sampleRowThreshold);
    }

    private SuggestionCacheEntry EnsureSuggestionCache(string sheetName, int columnIndex, int sampleRowThreshold)
    {
        var key = (sheetName, columnIndex, sampleRowThreshold);
        if (_suggestionCache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var worksheet = GetWorksheetOrThrow(sheetName);
        var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var endRow = worksheet.Dimension?.End.Row ?? 0;
        var effectiveThreshold = Math.Max(sampleRowThreshold > 0 ? sampleRowThreshold : 4, 4);

        if (columnIndex > 0 && endRow > effectiveThreshold)
        {
            for (var row = effectiveThreshold + 1; row <= endRow; row++)
            {
                var normalized = NormalizeSuggestionText(worksheet.Cells[row, columnIndex].Text ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(normalized))
                {
                    values.Add(normalized);
                }
            }

            // Trừ dòng mẫu trở lên: loại bỏ bất kỳ giá trị nào trùng với dòng mẫu hoặc các hàng tiêu đề phía trên
            for (var r = 1; r <= effectiveThreshold; r++)
            {
                var excludedText = NormalizeSuggestionText(worksheet.Cells[r, columnIndex].Text ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(excludedText))
                {
                    values.Remove(excludedText);
                }
            }
        }

        var entry = new SuggestionCacheEntry(values);
        _suggestionCache[key] = entry;
        return entry;
    }

    private bool ColumnContainsSuggestionValueOnSheet(string sheetName, int columnIndex, int sampleRowThreshold, string value)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var normalizedTarget = value.Trim();
        if (columnIndex <= 0 || string.IsNullOrWhiteSpace(normalizedTarget) || IsPureNumericText(normalizedTarget))
        {
            return false;
        }

        var endRow = worksheet.Dimension?.End.Row ?? 0;
        var effectiveThreshold = Math.Max(sampleRowThreshold > 0 ? sampleRowThreshold : 4, 4);
        if (endRow <= effectiveThreshold)
        {
            return false;
        }

        for (var row = effectiveThreshold + 1; row <= endRow; row++)
        {
            var text = worksheet.Cells[row, columnIndex].Text?.Trim() ?? string.Empty;
            if (string.Equals(text, normalizedTarget, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeSuggestionText(string text)
    {
        var normalized = text.Trim();
        if (string.IsNullOrWhiteSpace(normalized) || IsPureNumericText(normalized))
        {
            return string.Empty;
        }

        return normalized;
    }

    private string GetActiveSheetName()
    {
        EnsureWorksheet();
        return _worksheet!.Name;
    }

    public IReadOnlyList<string> GetDropdownOptions(int rowIndex, int columnIndex)
    {
        EnsureWorksheet();
        return GetDropdownOptions(_worksheet!.Name, rowIndex, columnIndex);
    }

    public IReadOnlyList<string> GetDropdownOptions(int rowIndex, int columnIndex, IReadOnlyDictionary<CellKey, string>? valueOverrides)
    {
        EnsureWorksheet();
        return GetDropdownOptions(_worksheet!.Name, rowIndex, columnIndex, valueOverrides);
    }

    public DropdownSearchIndex GetOrCreateDropdownSearchIndex(
        string sheetName,
        int rowIndex,
        int columnIndex,
        IReadOnlyDictionary<CellKey, string>? valueOverrides = null)
    {
        var options = GetDropdownOptions(sheetName, rowIndex, columnIndex, valueOverrides);
        return new DropdownSearchIndex(options);
    }

    public DropdownDependencyInfo GetDropdownDependencyInfo(string sheetName, int rowIndex, int columnIndex)
    {
        var extendedRule = _extendedValidations.FindRule(sheetName, rowIndex, columnIndex);
        if (extendedRule is not null && extendedRule.IsDynamic)
        {
            return new DropdownDependencyInfo(true, extendedRule.ParentColumns);
        }

        var standardFormula = TryGetStandardListValidationFormula(sheetName, rowIndex, columnIndex);
        if (!string.IsNullOrWhiteSpace(standardFormula)
            && ValidationFormulaAdjuster.LooksLikeDynamicFormula(standardFormula))
        {
            return new DropdownDependencyInfo(
                true,
                ValidationFormulaAdjuster.ExtractSameSheetParentColumns(standardFormula));
        }

        return DropdownDependencyInfo.Static;
    }

    public IReadOnlyList<string> GetDropdownOptions(
        string sheetName,
        int rowIndex,
        int columnIndex,
        IReadOnlyDictionary<CellKey, string>? valueOverrides = null)
    {
        EnsureWorksheet();
        var worksheet = GetWorksheetOrThrow(sheetName);
        var extendedRule = _extendedValidations.FindRule(sheetName, rowIndex, columnIndex);
        if (extendedRule is not null)
        {
            return ResolveExtendedDropdownOptions(worksheet, sheetName, rowIndex, columnIndex, extendedRule, valueOverrides);
        }

        var staticCacheKey = (sheetName, rowIndex, columnIndex);
        if (valueOverrides is null || valueOverrides.Count == 0)
        {
            if (_dropdownCellCache.TryGetValue(staticCacheKey, out var cached))
            {
                return cached;
            }

            var columnKey = (sheetName, columnIndex);
            if (_dropdownColumnCache.TryGetValue(columnKey, out var columnCached))
            {
                _dropdownCellCache[staticCacheKey] = columnCached;
                return columnCached;
            }
        }

        var resolved = ResolveStandardDropdownOptions(worksheet, sheetName, rowIndex, columnIndex);
        if (valueOverrides is null || valueOverrides.Count == 0)
        {
            CacheDropdownResults(staticCacheKey, (sheetName, columnIndex), resolved, isDynamic: false);
        }

        return resolved;
    }

    private IReadOnlyList<string> ResolveExtendedDropdownOptions(
        ExcelWorksheet worksheet,
        string sheetName,
        int rowIndex,
        int columnIndex,
        ExtendedListValidationRule rule,
        IReadOnlyDictionary<CellKey, string>? valueOverrides)
    {
        if (_package is null)
        {
            return [];
        }

        string? ReadCellValue(int targetRow, int targetColumn)
        {
            var key = new CellKey(sheetName, targetRow, targetColumn);
            if (valueOverrides is not null && valueOverrides.TryGetValue(key, out var overridden))
            {
                return overridden;
            }

            return worksheet.Cells[targetRow, targetColumn].Text;
        }

        return _dynamicDropdownResolver.Resolve(
            _package,
            worksheet,
            _workbookPath,
            sheetName,
            rowIndex,
            columnIndex,
            rule,
            ReadCellValue,
            formula => TryResolveDynamicListFormulaEpplus(formula));
    }

    private string? TryGetStandardListValidationFormula(string sheetName, int rowIndex, int columnIndex)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var validation = worksheet.DataValidations
            .OfType<IExcelDataValidation>()
            .FirstOrDefault(v => v.ValidationType.Type == eDataValidationType.List
                && IsCellInValidationAddress(v.Address?.Address ?? string.Empty, rowIndex, columnIndex));
        if (validation is not IExcelDataValidationList listValidation)
        {
            return null;
        }

        var inline = listValidation.Formula?.Values?
            .Select(v => v?.ToString()?.Trim() ?? string.Empty)
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        if (!string.IsNullOrWhiteSpace(inline))
        {
            return null;
        }

        return listValidation.Formula?.ExcelFormula?.Trim();
    }

    private IReadOnlyList<string> ResolveStandardDropdownOptions(
        ExcelWorksheet worksheet,
        string sheetName,
        int rowIndex,
        int columnIndex)
    {
        var validation = worksheet.DataValidations
            .OfType<IExcelDataValidation>()
            .FirstOrDefault(v => v.ValidationType.Type == eDataValidationType.List
                && IsCellInValidationAddress(v.Address?.Address ?? string.Empty, rowIndex, columnIndex));
        if (validation is not IExcelDataValidationList listValidation)
        {
            return [];
        }

        var values = listValidation.Formula?.Values?
            .Select(v => v?.ToString()?.Trim() ?? string.Empty)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList() ?? [];
        if (values.Count > 0)
        {
            return values;
        }

        var formula = listValidation.Formula?.ExcelFormula?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(formula))
        {
            return TryInteropDropdownFallback(worksheet, rowIndex, columnIndex, string.Empty);
        }

        IReadOnlyList<string> resolved;
        try
        {
            resolved = ResolveListFormulaToValues(formula, worksheet).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
        catch (Exception)
        {
            resolved = [];
        }

        if (resolved.Count > 0)
        {
            return resolved;
        }

        var dynamicFallback = TryResolveDynamicListFormulaEpplus(formula);
        if (dynamicFallback.Count > 0)
        {
            return dynamicFallback;
        }

        return TryInteropDropdownFallback(worksheet, rowIndex, columnIndex, formula);
    }

    private void CacheDropdownResults(
        (string Sheet, int Row, int Column) cellKey,
        (string Sheet, int Column) columnKey,
        IReadOnlyList<string> values,
        bool isDynamic)
    {
        _dropdownCellCache[cellKey] = values;
        if (!isDynamic && values.Count > 0)
        {
            _dropdownColumnCache[columnKey] = values;
        }
    }

    public void UpdateCellValue(int rowIndex, int columnIndex, string value)
    {
        EnsureWorksheet();
        UpdateCellValue(_worksheet!.Name, rowIndex, columnIndex, value);
    }

    public void UpdateCellValue(string sheetName, int rowIndex, int columnIndex, string value)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        worksheet.Cells[rowIndex, columnIndex].Value = value;
        _hasUnsavedChanges = true;
    }

    /// <summary>
    /// Ghi giá trị mà không đổi Style/Numberformat của ô đích.
    /// Cột text trên file gốc nhận đúng chuỗi hiển thị từ file số liệu; cột số/ngày chuyển kiểu theo dòng mẫu.
    /// </summary>
    public void UpdateCellValuePreservingFormat(int rowIndex, int columnIndex, LinkCellSnapshot source, int formatReferenceRow)
    {
        EnsureWorksheet();
        UpdateCellValuePreservingFormat(_worksheet!.Name, rowIndex, columnIndex, source, formatReferenceRow);
    }

    public void UpdateCellValuePreservingFormat(
        string sheetName,
        int rowIndex,
        int columnIndex,
        LinkCellSnapshot source,
        int formatReferenceRow)
    {
        var worksheet = GetWorksheetOrThrow(sheetName);
        var cell = worksheet.Cells[rowIndex, columnIndex];
        if (IsEmptySnapshot(source))
        {
            cell.Value = null;
            _hasUnsavedChanges = true;
            return;
        }

        var referenceCell = formatReferenceRow > 0
            ? worksheet.Cells[formatReferenceRow, columnIndex]
            : cell;
        cell.Value = ResolveLinkedCellValue(source, referenceCell);
        _hasUnsavedChanges = true;
    }

    private static bool IsEmptySnapshot(LinkCellSnapshot source)
    {
        if (!string.IsNullOrEmpty(source.DisplayText))
        {
            return false;
        }

        return source.RawValue is null
            || (source.RawValue is string rawText && string.IsNullOrWhiteSpace(rawText));
    }

    private static object? ResolveLinkedCellValue(LinkCellSnapshot source, ExcelRange referenceCell)
    {
        var numberFormat = referenceCell.Style.Numberformat.Format ?? string.Empty;
        if (IsTextTargetColumn(numberFormat, referenceCell))
        {
            return source.DisplayText;
        }

        if (source.RawValue is DateTime dateTime && IsDateTargetColumn(numberFormat, referenceCell))
        {
            return dateTime;
        }

        if (source.RawValue is double or decimal or int or long or float && IsNumericTargetColumn(numberFormat, referenceCell))
        {
            return Convert.ToDouble(source.RawValue, CultureInfo.InvariantCulture);
        }

        return CoerceTextToCellValue(source.DisplayText.Trim(), referenceCell);
    }

    private static bool IsTextTargetColumn(string numberFormat, ExcelRange referenceCell)
    {
        if (numberFormat.Contains('@', StringComparison.Ordinal))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(numberFormat)
            && !numberFormat.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var referenceValue = referenceCell.Value;
        return referenceValue is null or string
            || referenceValue is not double and not decimal and not int and not long and not float and not DateTime and not bool;
    }

    private static bool IsDateTargetColumn(string numberFormat, ExcelRange referenceCell)
    {
        return referenceCell.Value is DateTime || IsDateNumberFormat(numberFormat);
    }

    private static bool IsNumericTargetColumn(string numberFormat, ExcelRange referenceCell)
    {
        return referenceCell.Value is double or decimal or int or long or float || IsNumericNumberFormat(numberFormat);
    }

    private static object CoerceTextToCellValue(string textValue, ExcelRange referenceCell)
    {
        var referenceValue = referenceCell.Value;
        if (referenceValue is double or decimal or int or long or float)
        {
            if (TryParseNumber(textValue, out var number))
            {
                return number;
            }
        }

        if (referenceValue is DateTime)
        {
            if (TryParseDateTime(textValue, out var dateTime))
            {
                return dateTime;
            }
        }

        if (referenceValue is bool)
        {
            if (TryParseBoolean(textValue, out var boolean))
            {
                return boolean;
            }
        }

        var numberFormat = referenceCell.Style.Numberformat.Format ?? string.Empty;
        if (IsDateNumberFormat(numberFormat) && TryParseDateTime(textValue, out var parsedDate))
        {
            return parsedDate;
        }

        if (IsNumericNumberFormat(numberFormat) && TryParseNumber(textValue, out var parsedNumber))
        {
            return parsedNumber;
        }

        return textValue;
    }

    private static bool TryParseNumber(string text, out double number)
    {
        var normalized = text.Trim().Replace(" ", string.Empty);
        return double.TryParse(normalized, NumberStyles.Number, CultureInfo.CurrentCulture, out number)
            || double.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out number);
    }

    private static bool TryParseDateTime(string text, out DateTime dateTime)
    {
        return DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime)
            || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);
    }

    private static bool TryParseBoolean(string text, out bool value)
    {
        var normalized = text.Trim();
        if (bool.TryParse(normalized, out value))
        {
            return true;
        }

        if (string.Equals(normalized, "1", StringComparison.Ordinal) || string.Equals(normalized, "có", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (string.Equals(normalized, "0", StringComparison.Ordinal) || string.Equals(normalized, "không", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        return false;
    }

    private static bool IsDateNumberFormat(string format)
    {
        if (string.IsNullOrWhiteSpace(format) || format.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var lower = format.ToLowerInvariant();
        return lower.Contains('y')
            || lower.Contains("dd")
            || lower.Contains("mm")
            || lower.Contains("hh")
            || lower.Contains("ss");
    }

    private static bool IsNumericNumberFormat(string format)
    {
        if (string.IsNullOrWhiteSpace(format) || format.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return format.IndexOfAny(['#', '0', '%']) >= 0 && !IsDateNumberFormat(format);
    }

    public bool TrySave([NotNullWhen(false)] out string? userFacingMessage)
    {
        userFacingMessage = null;
        if (_package is null)
        {
            return true;
        }

        try
        {
            _package.Save();
            _hasUnsavedChanges = false;
            return true;
        }
        catch (Exception ex)
        {
            // EPPlus / IO có thể ném InvalidOperationException, TargetInvocationException, v.v. — không chỉ IOException.
            userFacingMessage = BuildSaveFailureMessage(ex);
            return false;
        }
    }

    public bool TrySaveAs(string path, [NotNullWhen(false)] out string? userFacingMessage)
    {
        userFacingMessage = null;
        if (_package is null)
        {
            userFacingMessage = "Chưa mở file Excel.";
            return false;
        }

        try
        {
            _package.SaveAs(new FileInfo(path));
            _workbookPath = path;
            _hasUnsavedChanges = false;
            return true;
        }
        catch (Exception ex)
        {
            userFacingMessage = BuildSaveFailureMessage(ex);
            return false;
        }
    }

    private static Exception GetInnermostException(Exception ex)
    {
        while (ex.InnerException is not null)
        {
            ex = ex.InnerException;
        }

        return ex;
    }

    private static string BuildSaveFailureMessage(Exception ex)
    {
        var root = GetInnermostException(ex);
        return "Không lưu được file. Có thể file đang được mở bởi Excel hoặc ứng dụng khác — hãy đóng file ở đó rồi thử lại.\n\n"
            + $"Chi tiết: {root.GetType().Name}: {root.Message}";
    }

    public bool HasUnsavedChanges() => _hasUnsavedChanges;

    public void CommitPendingEditsToWorksheetOnly()
    {
        _hasUnsavedChanges = true;
    }

    public void InvalidateActiveSheetSuggestionCaches()
    {
        EnsureWorksheet();
        var sheet = GetActiveSheetName();
        var keys = _suggestionCache.Keys.Where(k => string.Equals(k.Sheet, sheet, StringComparison.Ordinal)).ToList();
        foreach (var key in keys)
        {
            _suggestionCache.Remove(key);
        }
    }

    public void InvalidateActiveSheetDropdownCaches()
    {
        EnsureWorksheet();
        InvalidateSheetDropdownCaches(GetActiveSheetName());
    }

    private void InvalidateSheetDropdownCaches(string sheetName)
    {
        var keys = _dropdownCellCache.Keys.Where(k => string.Equals(k.Sheet, sheetName, StringComparison.Ordinal)).ToList();
        foreach (var key in keys)
        {
            _dropdownCellCache.Remove(key);
        }

        var columnKeys = _dropdownColumnCache.Keys.Where(k => string.Equals(k.Sheet, sheetName, StringComparison.Ordinal)).ToList();
        foreach (var key in columnKeys)
        {
            _dropdownColumnCache.Remove(key);
        }

        var indexKeys = _dropdownSearchIndexCache.Keys.Where(k => string.Equals(k.Sheet, sheetName, StringComparison.Ordinal)).ToList();
        foreach (var key in indexKeys)
        {
            _dropdownSearchIndexCache.Remove(key);
        }
    }

    public void InvalidateActiveSheetCaches()
    {
        InvalidateActiveSheetSuggestionCaches();
        InvalidateActiveSheetDropdownCaches();
    }

    public void InvalidateSheetCaches(string sheetName)
    {
        var suggestionKeys = _suggestionCache.Keys.Where(k => string.Equals(k.Sheet, sheetName, StringComparison.Ordinal)).ToList();
        foreach (var key in suggestionKeys)
        {
            _suggestionCache.Remove(key);
        }

        InvalidateSheetDropdownCaches(sheetName);
    }

    public void Dispose()
    {
        DisposePackageOnly();
        GC.SuppressFinalize(this);
    }

    private void EnsureWorksheet()
    {
        if (_worksheet is null)
        {
            throw new InvalidOperationException("Bạn cần chọn sheet trước.");
        }
    }

    private void EnsurePackage()
    {
        if (_package is null)
        {
            throw new InvalidOperationException("Chưa mở file Excel.");
        }
    }

    private ExcelWorksheet GetWorksheetOrThrow(string sheetName)
    {
        EnsurePackage();
        var worksheet = _package!.Workbook.Worksheets[sheetName];
        if (worksheet is null)
        {
            throw new InvalidOperationException($"Không tìm thấy sheet '{sheetName}'.");
        }

        return worksheet;
    }

    private IReadOnlyList<string> TryInteropDropdownFallback(
        ExcelWorksheet worksheet,
        int rowIndex,
        int columnIndex,
        string formula)
    {
        if (string.IsNullOrWhiteSpace(_workbookPath) || !ExcelInstallationProbe.IsInstalled())
        {
            return [];
        }

        try
        {
            return _interopDropdownReader.TryGetDropdownOptions(_workbookPath, worksheet.Name, rowIndex, columnIndex);
        }
        catch (FileNotFoundException)
        {
            ExcelInstallationProbe.MarkUnavailable();
            return TryResolveDynamicListFormulaEpplus(formula);
        }
        catch (Exception)
        {
            ExcelInstallationProbe.MarkUnavailable();
            return [];
        }
    }

    /// <summary>
    /// Fallback không cần Excel cài sẵn: đọc các cột được tham chiếu trong OFFSET/MATCH/INDIRECT.
    /// </summary>
    private IReadOnlyList<string> TryResolveDynamicListFormulaEpplus(string formula)
    {
        var normalized = formula.Trim();
        if (normalized.StartsWith('='))
        {
            normalized = normalized[1..];
        }

        if (!LooksLikeDynamicFormula(normalized))
        {
            return [];
        }

        var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in Regex.Matches(normalized, @"(?:'([^']+)'|([^'!]+))!\$?([A-Za-z]+)"))
        {
            var sheetName = !string.IsNullOrEmpty(match.Groups[1].Value)
                ? match.Groups[1].Value
                : match.Groups[2].Value.Trim();
            var columnIndex = ColumnLettersToIndex(match.Groups[3].Value);
            if (string.IsNullOrWhiteSpace(sheetName) || columnIndex < 1)
            {
                continue;
            }

            var targetWorksheet = TryGetWorksheet(sheetName);
            if (targetWorksheet is null)
            {
                continue;
            }

            var endRow = targetWorksheet.Dimension?.End.Row ?? 0;
            for (var row = 1; row <= endRow; row++)
            {
                var text = targetWorksheet.Cells[row, columnIndex].Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    values.Add(text);
                }
            }
        }

        return values
            .OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
    }

    private IReadOnlyList<string> ResolveListFormulaToValues(string excelFormula, ExcelWorksheet worksheet)
    {
        try
        {
            return ResolveListFormulaToValuesCore(excelFormula, worksheet);
        }
        catch (KeyNotFoundException)
        {
            return [];
        }
        catch (ArgumentException)
        {
            return [];
        }
        catch (InvalidOperationException)
        {
            return [];
        }
    }

    private IReadOnlyList<string> ResolveListFormulaToValuesCore(string excelFormula, ExcelWorksheet worksheet)
    {
        var normalized = excelFormula.Trim();
        if (normalized.StartsWith("=", StringComparison.Ordinal))
        {
            normalized = normalized[1..];
        }

        if (normalized.StartsWith("\"", StringComparison.Ordinal) && normalized.EndsWith("\"", StringComparison.Ordinal))
        {
            return normalized.Trim('"')
                .Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
        }

        // OFFSET/MATCH/INDIRECT and other dynamic formulas cannot be resolved statically via EPPlus.
        if (LooksLikeDynamicFormula(normalized))
        {
            return [];
        }

        // Some workbooks store list validation as plain comma/semicolon text without quotes.
        if ((normalized.Contains(',') || normalized.Contains(';')) && !normalized.Contains('!') && !normalized.Contains('['))
        {
            var inline = normalized
                .Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (inline.Count > 0)
            {
                return inline;
            }
        }

        var workbookNamedRange = TryGetWorkbookNamedRange(normalized);
        if (workbookNamedRange is not null)
        {
            return ReadNamedRangeValues(workbookNamedRange);
        }

        var worksheetNamedRange = TryGetWorksheetNamedRange(worksheet, normalized);
        if (worksheetNamedRange is not null)
        {
            return ReadNamedRangeValues(worksheetNamedRange);
        }

        var tableValues = ResolveTableReferenceToValues(normalized);
        if (tableValues.Count > 0)
        {
            return tableValues;
        }

        var (sheetName, rangeAddress) = SplitSheetAndRange(normalized);
        if (string.IsNullOrWhiteSpace(rangeAddress) || !IsSimpleRangeAddress(rangeAddress))
        {
            return [];
        }

        var targetWorksheet = string.IsNullOrWhiteSpace(sheetName)
            ? worksheet
            : TryGetWorksheet(sheetName);
        if (targetWorksheet is null)
        {
            return [];
        }

        var result = new List<string>();
        foreach (var cell in targetWorksheet.Cells[rangeAddress])
        {
            var text = cell.Text?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text))
            {
                result.Add(text);
            }
        }

        return result;
    }

    private static bool LooksLikeDynamicFormula(string normalized) =>
        ValidationFormulaAdjuster.LooksLikeDynamicFormula(normalized);

    private static bool IsSimpleRangeAddress(string rangeAddress) =>
        Regex.IsMatch(rangeAddress, @"^[A-Za-z]+\d+(?::[A-Za-z]+\d+)?$")
        || Regex.IsMatch(rangeAddress, @"^[A-Za-z]+:[A-Za-z]+$");

    private ExcelNamedRange? TryGetWorkbookNamedRange(string name)
    {
        if (_package is null || string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        return _package.Workbook.Names
            .FirstOrDefault(n => string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private static ExcelNamedRange? TryGetWorksheetNamedRange(ExcelWorksheet worksheet, string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        return worksheet.Names
            .FirstOrDefault(n => string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private ExcelWorksheet? TryGetWorksheet(string sheetName)
    {
        if (_package is null || string.IsNullOrWhiteSpace(sheetName))
        {
            return null;
        }

        return _package.Workbook.Worksheets
            .FirstOrDefault(ws => string.Equals(ws.Name, sheetName, StringComparison.Ordinal));
    }

    private static List<string> ReadNamedRangeValues(ExcelNamedRange namedRange) =>
        namedRange
            .Select(c => c.Text?.Trim() ?? string.Empty)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private IReadOnlyList<string> ResolveTableReferenceToValues(string formula)
    {
        if (_package is null)
        {
            return [];
        }

        var match = Regex.Match(formula, @"^(?<table>[^\[]+)\[(?<column>.+)\]$");
        if (!match.Success)
        {
            return [];
        }

        var tableName = match.Groups["table"].Value.Trim();
        var columnToken = match.Groups["column"].Value.Trim();
        columnToken = columnToken.Replace("[", string.Empty).Replace("]", string.Empty);
        var columnName = columnToken
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .LastOrDefault(token => !token.StartsWith("#", StringComparison.Ordinal)) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(columnName))
        {
            return [];
        }

        foreach (var worksheet in _package.Workbook.Worksheets)
        {
            var table = worksheet.Tables.FirstOrDefault(t => string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase));
            if (table is null)
            {
                continue;
            }

            var columnIndex = FindTableColumnOffset(table, columnName);
            if (columnIndex < 0)
            {
                return [];
            }

            var absoluteColumn = table.Address.Start.Column + columnIndex;
            var startRow = table.Address.Start.Row + 1;
            var endRow = table.Address.End.Row;
            var values = new List<string>();
            for (var row = startRow; row <= endRow; row++)
            {
                var text = worksheet.Cells[row, absoluteColumn].Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    values.Add(text);
                }
            }

            return values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        return [];
    }

    private static int FindTableColumnOffset(ExcelTable table, string columnName)
    {
        for (var i = 0; i < table.Columns.Count; i++)
        {
            if (string.Equals(table.Columns[i].Name, columnName, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    private static (string SheetName, string RangeAddress) SplitSheetAndRange(string formula)
    {
        var exclamationIndex = formula.LastIndexOf('!');
        if (exclamationIndex < 0)
        {
            return (string.Empty, formula);
        }

        var sheetPart = formula[..exclamationIndex].Trim().Trim('\'');
        var rangePart = formula[(exclamationIndex + 1)..].Trim();
        return (sheetPart, rangePart);
    }

    private static bool IsCellInValidationAddress(string addressText, int rowIndex, int columnIndex)
    {
        var cleaned = addressText.Replace("$", string.Empty);
        var ranges = cleaned.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var rawRange in ranges)
        {
            var range = rawRange;
            var excl = range.LastIndexOf('!');
            if (excl >= 0)
            {
                range = range[(excl + 1)..];
            }

            if (IsCellInRange(range, rowIndex, columnIndex))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsCellInRange(string range, int rowIndex, int columnIndex)
    {
        var pureColumnRange = Regex.Match(range, @"^([A-Za-z]+):([A-Za-z]+)$");
        if (pureColumnRange.Success)
        {
            var startColOnly = ColumnLettersToIndex(pureColumnRange.Groups[1].Value);
            var endColOnly = ColumnLettersToIndex(pureColumnRange.Groups[2].Value);
            var colMinOnly = Math.Min(startColOnly, endColOnly);
            var colMaxOnly = Math.Max(startColOnly, endColOnly);
            return columnIndex >= colMinOnly && columnIndex <= colMaxOnly;
        }

        var pureRowRange = Regex.Match(range, @"^(\d+):(\d+)$");
        if (pureRowRange.Success
            && int.TryParse(pureRowRange.Groups[1].Value, out var startRowOnly)
            && int.TryParse(pureRowRange.Groups[2].Value, out var endRowOnly))
        {
            var rowMinOnly = Math.Min(startRowOnly, endRowOnly);
            var rowMaxOnly = Math.Max(startRowOnly, endRowOnly);
            return rowIndex >= rowMinOnly && rowIndex <= rowMaxOnly;
        }

        var parts = range.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 1)
        {
            return TryParseCell(parts[0], out var row, out var col) && row == rowIndex && col == columnIndex;
        }

        if (!TryParseCell(parts[0], out var startRow, out var startCol) || !TryParseCell(parts[1], out var endRow, out var endCol))
        {
            return false;
        }

        var rowMin = Math.Min(startRow, endRow);
        var rowMax = Math.Max(startRow, endRow);
        var colMin = Math.Min(startCol, endCol);
        var colMax = Math.Max(startCol, endCol);
        return rowIndex >= rowMin && rowIndex <= rowMax && columnIndex >= colMin && columnIndex <= colMax;
    }

    private static bool TryParseCell(string cellAddress, out int row, out int column)
    {
        row = 0;
        column = 0;
        var pureColumn = Regex.Match(cellAddress, @"^([A-Za-z]+)$");
        if (pureColumn.Success)
        {
            row = 1;
            column = ColumnLettersToIndex(pureColumn.Groups[1].Value);
            return true;
        }

        var pureRow = Regex.Match(cellAddress, @"^(\d+)$");
        if (pureRow.Success && int.TryParse(pureRow.Groups[1].Value, out var parsedRow))
        {
            row = parsedRow;
            column = 1;
            return true;
        }

        var match = Regex.Match(cellAddress, @"^([A-Za-z]+)(\d+)$");
        if (!match.Success)
        {
            return false;
        }

        column = ColumnLettersToIndex(match.Groups[1].Value);
        return int.TryParse(match.Groups[2].Value, out row);
    }

    private static int ColumnLettersToIndex(string letters)
    {
        var result = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            result = (result * 26) + (ch - 'A' + 1);
        }

        return result;
    }

    public void Close()
    {
        DisposePackageOnly();
    }

    public void DisposePackageOnly()
    {
        _worksheet = null;
        _package?.Dispose();
        _package = null;
        _suggestionCache.Clear();
        _dropdownCellCache.Clear();
        _dropdownColumnCache.Clear();
        _dropdownSearchIndexCache.Clear();
    }
}

internal sealed class SuggestionCacheEntry(HashSet<string> values)
{
    public HashSet<string> Values { get; } = values;
}

public sealed record WorksheetInfo(string Name, bool IsHidden);
public sealed record HeaderCell(int ColumnIndex, string Name);
public sealed record RecordRow(int RowIndex, string KeyDisplay, string TooltipPreview);
public sealed record LogicalRecordRow(int LogicalIndex, int RowIndex, string KeyDisplay, string TooltipPreview);
