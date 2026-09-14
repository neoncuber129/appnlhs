using OfficeOpenXml;
using System.IO;

namespace ExcelDataEntryApp.Services;

public sealed class ExcelLookupReader : IDisposable
{
    private ExcelPackage? _package;
    private ExcelWorksheet? _worksheet;

    public void Open(string path, string sheetName)
    {
        DisposePackage();
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Không tìm thấy file số liệu.", path);
        }

        ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
        _package = new ExcelPackage(new FileInfo(path));
        _worksheet = _package.Workbook.Worksheets[sheetName]
            ?? throw new InvalidOperationException($"Không tìm thấy sheet '{sheetName}' trong file số liệu.");
    }

    public IReadOnlyList<string> GetWorksheetNames()
    {
        return _package?.Workbook.Worksheets.Select(s => s.Name).ToList() ?? [];
    }

    public IReadOnlyList<HeaderCell> ReadHeaderRow(int headerRow)
    {
        EnsureWorksheet();
        var endCol = _worksheet!.Dimension?.End.Column ?? 0;
        var headers = new List<HeaderCell>();
        for (var col = 1; col <= endCol; col++)
        {
            var text = _worksheet.Cells[headerRow, col].Text?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                text = $"Column {col}";
            }

            headers.Add(new HeaderCell(col, text));
        }

        return headers;
    }

    public string ReadCellText(int rowIndex, int columnIndex)
    {
        EnsureWorksheet();
        return _worksheet!.Cells[rowIndex, columnIndex].Text?.Trim() ?? string.Empty;
    }

    public LinkCellSnapshot ReadCellSnapshot(int rowIndex, int columnIndex)
    {
        EnsureWorksheet();
        var cell = _worksheet!.Cells[rowIndex, columnIndex];
        return new LinkCellSnapshot
        {
            DisplayText = cell.Text ?? string.Empty,
            RawValue = cell.Value
        };
    }

    public int GetEndRow()
    {
        EnsureWorksheet();
        return _worksheet!.Dimension?.End.Row ?? 0;
    }

    public IReadOnlyList<(int RowIndex, string Key)> ScanKeyColumn(int firstDataRow, int keyColumnIndex)
    {
        EnsureWorksheet();
        var endRow = GetEndRow();
        var rows = new List<(int RowIndex, string Key)>();
        for (var row = firstDataRow; row <= endRow; row++)
        {
            var key = ReadCellText(row, keyColumnIndex);
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            rows.Add((row, key));
        }

        return rows;
    }

    public void Dispose()
    {
        DisposePackage();
    }

    private void EnsureWorksheet()
    {
        if (_worksheet is null)
        {
            throw new InvalidOperationException("Chưa mở file số liệu.");
        }
    }

    private void DisposePackage()
    {
        _worksheet = null;
        _package?.Dispose();
        _package = null;
    }
}
