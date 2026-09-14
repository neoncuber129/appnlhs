namespace ExcelDataEntryApp.Models;

public sealed class HsskExportProfile
{
    public int GenderColumnIndex { get; set; }

    public List<int> SkipSampleForMaleColumnIndexes { get; set; } = [];
}
