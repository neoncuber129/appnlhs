namespace ExcelDataEntryApp.Models;

/// <summary>Cấu hình import danh sách: cột tên trên file danh sách + ánh xạ sang file gốc.</summary>
public sealed class ImportListProfile
{
    public string DataFilePath { get; set; } = string.Empty;

    public string DataSheetName { get; set; } = string.Empty;

    public int DataHeaderRow { get; set; } = 1;

    /// <summary>Cột tên trên file danh sách (dùng để chọn người import).</summary>
    public int DataNameColumnIndex { get; set; }

    public List<ColumnMappingPair> ColumnMappings { get; set; } = [];
}
