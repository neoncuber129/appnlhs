namespace ExcelDataEntryApp.Models;

public sealed class NamedDataLinkProfile
{
    public string Name { get; set; } = string.Empty;
    public DataLinkProfile Profile { get; set; } = new();
}
