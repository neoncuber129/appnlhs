using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.IO;

namespace ExcelDataEntryApp.Services;

public sealed class HeaderProfileStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _profilePath;

    public HeaderProfileStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "ExcelDataEntryApp");
        Directory.CreateDirectory(directory);
        _profilePath = Path.Combine(directory, "header-profiles.json");
    }

    public string BuildKey(string workbookPath, string sheetName, IEnumerable<string> headers)
    {
        var fingerprintInput = $"{Path.GetFileName(workbookPath)}|{sheetName}|{string.Join("|", headers)}";
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(fingerprintInput));
        return Convert.ToHexString(bytes);
    }

    public IReadOnlySet<string>? TryLoadHiddenHeaders(string key)
    {
        var data = LoadAll();
        if (data.TryGetValue(key, out var values))
        {
            return values.ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        return null;
    }

    public void SaveHiddenHeaders(string key, IEnumerable<string> hiddenHeaders)
    {
        var data = LoadAll();
        data[key] = hiddenHeaders.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_profilePath, json);
    }

    private Dictionary<string, List<string>> LoadAll()
    {
        if (!File.Exists(_profilePath))
        {
            return new Dictionary<string, List<string>>();
        }

        var json = File.ReadAllText(_profilePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new Dictionary<string, List<string>>();
        }

        return JsonSerializer.Deserialize<Dictionary<string, List<string>>>(json) ?? new Dictionary<string, List<string>>();
    }
}
