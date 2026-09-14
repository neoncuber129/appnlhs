namespace ExcelDataEntryApp.Models;

public sealed class ColumnMappingPair
{
    public int TargetColumnIndex { get; set; }

    public int SourceColumnIndex { get; set; }

    /// <summary>Sheet đích khi ghép nhiều sheet (rỗng = sheet hiện tại / một sheet).</summary>
    public string TargetSheetName { get; set; } = string.Empty;
}

public sealed class DataLinkProfile
{
    public string DataFilePath { get; set; } = string.Empty;
    public string DataSheetName { get; set; } = string.Empty;
    public int DataHeaderRow { get; set; } = 1;
    public int MainKeyColumnIndex { get; set; }

    /// <summary>Sheet chứa cột key file gốc (chế độ nhiều sheet).</summary>
    public string MainKeySheetName { get; set; } = string.Empty;

    public int DataKeyColumnIndex { get; set; }
    public List<ColumnMappingPair> ColumnMappings { get; set; } = [];

    /// <summary>Tự tạo dòng mới trên file gốc khi import danh sách (tên chưa có).</summary>
    public bool CreateMissingRowsInMain { get; set; }
}
