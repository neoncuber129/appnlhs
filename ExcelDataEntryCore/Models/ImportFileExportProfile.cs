namespace ExcelDataEntryApp.Models;

public sealed class ImportFileExportProfile
{
    public int GenderColumnIndex { get; set; }

    public List<int> SkipSampleColumnIndexes { get; set; } = [];
}
