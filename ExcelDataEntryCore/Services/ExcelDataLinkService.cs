using System.IO;
using ExcelDataEntryApp.Infrastructure;
using ExcelDataEntryApp.Models;
using OfficeOpenXml;

namespace ExcelDataEntryApp.Services;

public sealed class ExcelDataLinkService
{
    public DataLinkAnalysis AnalyzeKeyIssues(
        ExcelWorkbookService mainWorkbook,
        ExcelLookupReader dataReader,
        DataLinkMainContext mainContext,
        int dataFirstDataRow,
        int mainKeyColumnIndex,
        int dataKeyColumnIndex)
    {
        var dataRows = dataReader.ScanKeyColumn(dataFirstDataRow, dataKeyColumnIndex);
        var mainKeyRows = CollectMainKeyRows(mainWorkbook, mainContext, mainKeyColumnIndex);

        var mainDuplicates = FindDuplicates(mainKeyRows);
        var dataDuplicates = FindDuplicates(dataRows);

        var mainNormalizedKeys = mainKeyRows
            .Select(r => NormalizeKey(r.Key))
            .ToHashSet(StringComparer.Ordinal);
        var dataNormalizedKeys = dataRows
            .Select(r => NormalizeKey(r.Key))
            .ToHashSet(StringComparer.Ordinal);

        var keysOnlyInMain = mainKeyRows
            .GroupBy(r => NormalizeKey(r.Key), StringComparer.Ordinal)
            .Where(g => !dataNormalizedKeys.Contains(g.Key))
            .Select(g => g.First().Key)
            .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var keysOnlyInData = dataRows
            .GroupBy(r => NormalizeKey(r.Key), StringComparer.Ordinal)
            .Where(g => !mainNormalizedKeys.Contains(g.Key))
            .Select(g => g.First().Key)
            .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new DataLinkAnalysis
        {
            MainFileDuplicates = mainDuplicates,
            DataFileDuplicates = dataDuplicates,
            KeysOnlyInMain = keysOnlyInMain,
            KeysOnlyInData = keysOnlyInData
        };
    }

    public DataLinkResult RunLink(
        ExcelWorkbookService mainWorkbook,
        ExcelLookupReader dataReader,
        DataLinkMainContext mainContext,
        int dataFirstDataRow,
        int mainKeyColumnIndex,
        int dataKeyColumnIndex,
        IReadOnlyList<ColumnMappingPair> columnMappings,
        int singleSheetFormatReferenceRow,
        IReadOnlySet<string>? selectedDataNormalizedKeys = null)
    {
        var analysis = AnalyzeKeyIssues(
            mainWorkbook,
            dataReader,
            mainContext,
            dataFirstDataRow,
            mainKeyColumnIndex,
            dataKeyColumnIndex);

        var mainDuplicateKeys = BuildDuplicateKeySet(analysis.MainFileDuplicates);
        var dataDuplicateKeys = BuildDuplicateKeySet(analysis.DataFileDuplicates);

        var dataRows = dataReader.ScanKeyColumn(dataFirstDataRow, dataKeyColumnIndex);
        var lookupIndex = BuildLookupIndex(dataRows, dataDuplicateKeys);
        var mainKeyRows = CollectMainKeyRows(mainWorkbook, mainContext, mainKeyColumnIndex);

        var skippedRows = new List<SkippedLinkEntry>();
        var rowsMatched = 0;
        var cellsWritten = 0;

        foreach (var (rowIndex, key) in mainKeyRows)
        {
            var normalizedKey = NormalizeKey(key);
            if (selectedDataNormalizedKeys is not null
                && !selectedDataNormalizedKeys.Contains(normalizedKey))
            {
                continue;
            }

            if (mainDuplicateKeys.Contains(normalizedKey))
            {
                skippedRows.Add(new SkippedLinkEntry
                {
                    RowIndex = rowIndex,
                    KeyValue = key,
                    Reason = SkippedLinkReason.MainDuplicateKey
                });
                continue;
            }

            if (dataDuplicateKeys.Contains(normalizedKey))
            {
                skippedRows.Add(new SkippedLinkEntry
                {
                    RowIndex = rowIndex,
                    KeyValue = key,
                    Reason = SkippedLinkReason.DataDuplicateKey
                });
                continue;
            }

            if (!lookupIndex.TryGetValue(normalizedKey, out var sourceRow))
            {
                skippedRows.Add(new SkippedLinkEntry
                {
                    RowIndex = rowIndex,
                    KeyValue = key,
                    Reason = SkippedLinkReason.NoMatchInDataFile
                });
                continue;
            }

            rowsMatched++;
            foreach (var mapping in columnMappings)
            {
                var sourceCell = dataReader.ReadCellSnapshot(sourceRow, mapping.SourceColumnIndex);
                WriteMappedCell(
                    mainWorkbook,
                    mainContext,
                    rowIndex,
                    mapping,
                    sourceCell,
                    singleSheetFormatReferenceRow,
                    ref cellsWritten);
            }
        }

        return new DataLinkResult
        {
            RowsMatched = rowsMatched,
            CellsWritten = cellsWritten,
            MainFileDuplicates = analysis.MainFileDuplicates,
            DataFileDuplicates = analysis.DataFileDuplicates,
            KeysOnlyInMain = analysis.KeysOnlyInMain,
            KeysOnlyInData = analysis.KeysOnlyInData,
            SkippedRows = skippedRows
        };
    }

    public void ExportReport(string outputPath, DataLinkAnalysis analysis, DataLinkResult? linkResult = null)
    {
        var fileInfo = new FileInfo(outputPath);
        if (fileInfo.Directory is not null)
        {
            Directory.CreateDirectory(fileInfo.Directory.FullName);
        }

        ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
        using var package = new ExcelPackage();

        WriteDuplicateSheet(package, "Trung_key_file_goc", analysis.MainFileDuplicates);
        WriteDuplicateSheet(package, "Trung_key_file_so_lieu", analysis.DataFileDuplicates);
        WriteKeyListSheet(package, "Key_chi_co_file_goc", analysis.KeysOnlyInMain);
        WriteKeyListSheet(package, "Key_chi_co_file_so_lieu", analysis.KeysOnlyInData);

        if (linkResult is not null && linkResult.SkippedRows.Count > 0)
        {
            WriteSkippedRowsSheet(package, "Bo_qua_khi_lien_ket", linkResult.SkippedRows);
        }

        package.SaveAs(fileInfo);
    }

    private static HashSet<string> BuildDuplicateKeySet(IReadOnlyList<DuplicateKeyEntry> duplicates)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in duplicates)
        {
            set.Add(NormalizeKey(entry.KeyValue));
        }

        return set;
    }

    public int WriteImportRowFromList(
        ExcelWorkbookService mainWorkbook,
        ExcelLookupReader dataReader,
        DataLinkMainContext mainContext,
        int dataRowIndex,
        int mainKeyRowIndex,
        IReadOnlyList<ColumnMappingPair> columnMappings,
        int singleSheetFormatReferenceRow)
    {
        var cellsWritten = 0;
        foreach (var mapping in columnMappings)
        {
            var sourceCell = dataReader.ReadCellSnapshot(dataRowIndex, mapping.SourceColumnIndex);
            WriteMappedCell(
                mainWorkbook,
                mainContext,
                mainKeyRowIndex,
                mapping,
                sourceCell,
                singleSheetFormatReferenceRow,
                ref cellsWritten);
        }

        return cellsWritten;
    }

    private static void WriteMappedCell(
        ExcelWorkbookService mainWorkbook,
        DataLinkMainContext mainContext,
        int keyRowIndex,
        ColumnMappingPair mapping,
        LinkCellSnapshot sourceCell,
        int singleSheetFormatReferenceRow,
        ref int cellsWritten)
    {
        if (!mainContext.IsMultiSheet)
        {
            mainWorkbook.UpdateCellValuePreservingFormat(
                keyRowIndex,
                mapping.TargetColumnIndex,
                sourceCell,
                singleSheetFormatReferenceRow);
            cellsWritten++;
            return;
        }

        var targetSheet = string.IsNullOrEmpty(mapping.TargetSheetName)
            ? mainContext.KeySheetName
            : mapping.TargetSheetName;
        if (!mainContext.FirstDataRowBySheet.TryGetValue(targetSheet, out var targetFirstDataRow))
        {
            return;
        }

        var targetRow = string.Equals(targetSheet, mainContext.KeySheetName, StringComparison.Ordinal)
            ? keyRowIndex
            : targetFirstDataRow + (keyRowIndex - mainContext.KeySheetFirstDataRow);
        var formatRef = mainContext.FormatReferenceRowBySheet.TryGetValue(targetSheet, out var sampleRow)
            ? sampleRow
            : mainContext.KeySheetFirstDataRow;

        mainWorkbook.UpdateCellValuePreservingFormat(
            targetSheet,
            targetRow,
            mapping.TargetColumnIndex,
            sourceCell,
            formatRef);
        cellsWritten++;
    }

    private static List<(int RowIndex, string Key)> CollectMainKeyRows(
        ExcelWorkbookService mainWorkbook,
        DataLinkMainContext mainContext,
        int mainKeyColumnIndex)
    {
        var mainEndRow = string.IsNullOrEmpty(mainContext.KeySheetName)
            ? mainWorkbook.GetEndRow()
            : mainWorkbook.GetEndRow(mainContext.KeySheetName);
        var mainKeyRows = new List<(int RowIndex, string Key)>();
        for (var row = mainContext.KeySheetFirstDataRow; row <= mainEndRow; row++)
        {
            var key = string.IsNullOrEmpty(mainContext.KeySheetName)
                ? mainWorkbook.ReadCellText(row, mainKeyColumnIndex)
                : mainWorkbook.ReadCellText(mainContext.KeySheetName, row, mainKeyColumnIndex);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            mainKeyRows.Add((row, key));
        }

        return mainKeyRows;
    }

    private static Dictionary<string, int> BuildLookupIndex(
        IReadOnlyList<(int RowIndex, string Key)> rows,
        IReadOnlySet<string> duplicateKeys)
    {
        var index = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (rowIndex, key) in rows)
        {
            var normalized = NormalizeKey(key);
            if (duplicateKeys.Contains(normalized))
            {
                continue;
            }

            if (!index.ContainsKey(normalized))
            {
                index[normalized] = rowIndex;
            }
        }

        return index;
    }

    private static IReadOnlyList<DuplicateKeyEntry> FindDuplicates(IReadOnlyList<(int RowIndex, string Key)> rows)
    {
        var groups = rows
            .GroupBy(r => NormalizeKey(r.Key), StringComparer.Ordinal)
            .Where(g => g.Count() > 1)
            .OrderBy(g => g.Key, StringComparer.Ordinal)
            .Select(g =>
            {
                var variants = g.Select(x => x.Key)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return new DuplicateKeyEntry
                {
                    KeyValue = variants[0],
                    OccurrenceCount = g.Count(),
                    RowNumbers = g.Select(x => x.RowIndex).OrderBy(x => x).ToList(),
                    SpellingVariants = variants.Count > 1 ? variants : []
                };
            })
            .ToList();

        return groups;
    }

    private static string NormalizeKey(string key) => VietnameseTextHelper.ToTonePlacementKey(key);

    private static void WriteDuplicateSheet(ExcelPackage package, string sheetName, IReadOnlyList<DuplicateKeyEntry> duplicates)
    {
        var sheet = package.Workbook.Worksheets.Add(sheetName);
        sheet.Cells[1, 1].Value = "Giá trị key";
        sheet.Cells[1, 2].Value = "Số lần";
        sheet.Cells[1, 3].Value = "Các dòng";
        sheet.Cells[1, 4].Value = "Các biến thể ghi";

        var row = 2;
        foreach (var entry in duplicates)
        {
            sheet.Cells[row, 1].Value = entry.KeyValue;
            sheet.Cells[row, 2].Value = entry.OccurrenceCount;
            sheet.Cells[row, 3].Value = string.Join(", ", entry.RowNumbers);
            sheet.Cells[row, 4].Value = entry.SpellingVariants.Count > 0
                ? string.Join(" | ", entry.SpellingVariants)
                : string.Empty;
            row++;
        }

        if (sheet.Dimension is not null)
        {
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static void WriteKeyListSheet(ExcelPackage package, string sheetName, IReadOnlyList<string> keys)
    {
        var sheet = package.Workbook.Worksheets.Add(sheetName);
        sheet.Cells[1, 1].Value = "Giá trị key";

        var row = 2;
        foreach (var key in keys)
        {
            sheet.Cells[row, 1].Value = key;
            row++;
        }

        if (sheet.Dimension is not null)
        {
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static void WriteSkippedRowsSheet(ExcelPackage package, string sheetName, IReadOnlyList<SkippedLinkEntry> skippedRows)
    {
        var sheet = package.Workbook.Worksheets.Add(sheetName);
        sheet.Cells[1, 1].Value = "Dòng file gốc";
        sheet.Cells[1, 2].Value = "Giá trị key";
        sheet.Cells[1, 3].Value = "Lý do bỏ qua";

        var row = 2;
        foreach (var entry in skippedRows.OrderBy(e => e.RowIndex))
        {
            sheet.Cells[row, 1].Value = entry.RowIndex;
            sheet.Cells[row, 2].Value = entry.KeyValue;
            sheet.Cells[row, 3].Value = DescribeSkippedReason(entry.Reason);
            row++;
        }

        if (sheet.Dimension is not null)
        {
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }
    }

    private static string DescribeSkippedReason(SkippedLinkReason reason) => reason switch
    {
        SkippedLinkReason.MainDuplicateKey => "Key trùng trong file gốc",
        SkippedLinkReason.DataDuplicateKey => "Key trùng trong file số liệu",
        SkippedLinkReason.NoMatchInDataFile => "Không có số liệu khớp",
        _ => reason.ToString()
    };
}
