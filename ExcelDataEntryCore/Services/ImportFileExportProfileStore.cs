using System.IO;
using System.Text.Json;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp.Services;

public sealed class ImportFileExportProfileStore
{
    public const string FileName = "import-file-export-profiles.json";

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _legacyProfilePath;

    public ImportFileExportProfileStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        _legacyProfilePath = Path.Combine(appData, "ExcelDataEntryApp", FileName);
        TryMigrateFromLegacy();
    }

    public string ProfileFilePath => Path.Combine(AppContext.BaseDirectory, FileName);

    public ImportFileExportProfile? TryLoad(string profileKey)
    {
        var data = LoadAll();
        return data.TryGetValue(profileKey, out var profile) ? profile : null;
    }

    public void Save(string profileKey, ImportFileExportProfile profile)
    {
        var data = LoadAll();
        data[profileKey] = profile;
        WriteAll(data);
    }

    public Dictionary<string, ImportFileExportProfile> LoadAll()
    {
        if (!File.Exists(ProfileFilePath))
        {
            return new Dictionary<string, ImportFileExportProfile>();
        }

        var json = File.ReadAllText(ProfileFilePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new Dictionary<string, ImportFileExportProfile>();
        }

        return JsonSerializer.Deserialize<Dictionary<string, ImportFileExportProfile>>(json)
            ?? new Dictionary<string, ImportFileExportProfile>();
    }

    public void ReplaceAll(IReadOnlyDictionary<string, ImportFileExportProfile> profiles)
    {
        WriteAll(profiles.ToDictionary(static p => p.Key, static p => p.Value));
    }

    private void WriteAll(Dictionary<string, ImportFileExportProfile> data)
    {
        var directory = Path.GetDirectoryName(ProfileFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(ProfileFilePath, json);
    }

    private void TryMigrateFromLegacy()
    {
        if (File.Exists(ProfileFilePath) || !File.Exists(_legacyProfilePath))
        {
            return;
        }

        try
        {
            File.Copy(_legacyProfilePath, ProfileFilePath);
        }
        catch
        {
            // Bỏ qua nếu không copy được.
        }
    }
}
