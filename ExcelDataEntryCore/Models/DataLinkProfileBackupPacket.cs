namespace ExcelDataEntryApp.Models;

public sealed class DataLinkProfileBackupPacket
{
    public int SchemaVersion { get; set; } = 1;
    public string ProfileKey { get; set; } = string.Empty;
    public string? LastNamedProfile { get; set; }
    public List<NamedDataLinkProfile> Profiles { get; set; } = [];
}
