using System.IO;
using System.Text.Json;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp.Services;

public sealed class DataLinkProfileStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _profilePath;
    private readonly string _namedProfilePath;
    private readonly string _lastNamedProfilePath;

    public DataLinkProfileStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "ExcelDataEntryApp");
        Directory.CreateDirectory(directory);
        _profilePath = Path.Combine(directory, "data-link-profiles.json");
        _namedProfilePath = Path.Combine(directory, "data-link-named-profiles.json");
        _lastNamedProfilePath = Path.Combine(directory, "data-link-last-named.json");
    }

    public DataLinkProfile? TryLoad(string profileKey)
    {
        var data = LoadAll();
        if (data.TryGetValue(profileKey, out var profile))
        {
            return profile;
        }

        return null;
    }

    public void Save(string profileKey, DataLinkProfile profile)
    {
        var data = LoadAll();
        data[profileKey] = profile;
        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_profilePath, json);
    }

    public IReadOnlyList<string> GetSavedProfileNames(string profileKey)
    {
        var data = LoadNamedAll();
        if (!data.TryGetValue(profileKey, out var profiles))
        {
            return [];
        }

        return profiles.Select(p => p.Name).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public DataLinkProfile? TryLoadNamed(string profileKey, string name)
    {
        var data = LoadNamedAll();
        if (!data.TryGetValue(profileKey, out var profiles))
        {
            return null;
        }

        return profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))?.Profile;
    }

    public void SaveNamed(string profileKey, string name, DataLinkProfile profile)
    {
        var data = LoadNamedAll();
        if (!data.TryGetValue(profileKey, out var profiles))
        {
            profiles = [];
            data[profileKey] = profiles;
        }

        var existing = profiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            existing.Profile = profile;
        }
        else
        {
            profiles.Add(new NamedDataLinkProfile { Name = name, Profile = profile });
        }

        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_namedProfilePath, json);
        Save(profileKey, profile);
    }

    public string? TryLoadLastNamedProfileName(string profileKey)
    {
        var data = LoadLastNamedAll();
        return data.TryGetValue(profileKey, out var name) ? name : null;
    }

    public void SaveLastNamedProfileName(string profileKey, string name)
    {
        var data = LoadLastNamedAll();
        data[profileKey] = name;
        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_lastNamedProfilePath, json);
    }

    public void ClearLastNamedProfileName(string profileKey)
    {
        var data = LoadLastNamedAll();
        if (!data.Remove(profileKey))
        {
            return;
        }

        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_lastNamedProfilePath, json);
    }

    public void DeleteNamed(string profileKey, string name)
    {
        var data = LoadNamedAll();
        if (!data.TryGetValue(profileKey, out var profiles))
        {
            return;
        }

        profiles.RemoveAll(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        if (profiles.Count == 0)
        {
            data.Remove(profileKey);
        }

        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_namedProfilePath, json);
    }

    public void BackupNamedProfiles(string profileKey, string outputPath)
    {
        var namedData = LoadNamedAll();
        var profiles = namedData.TryGetValue(profileKey, out var existing)
            ? existing.Select(p => new NamedDataLinkProfile
            {
                Name = p.Name,
                Profile = p.Profile
            }).ToList()
            : [];
        var packet = new DataLinkProfileBackupPacket
        {
            ProfileKey = profileKey,
            LastNamedProfile = TryLoadLastNamedProfileName(profileKey),
            Profiles = profiles
        };

        var fileInfo = new FileInfo(outputPath);
        if (fileInfo.Directory is not null)
        {
            Directory.CreateDirectory(fileInfo.Directory.FullName);
        }

        var json = JsonSerializer.Serialize(packet, JsonOptions);
        File.WriteAllText(outputPath, json);
    }

    public IReadOnlyList<string> RestoreNamedProfiles(string profileKey, string inputPath)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Không tìm thấy file backup profile.", inputPath);
        }

        var json = File.ReadAllText(inputPath);
        if (string.IsNullOrWhiteSpace(json))
        {
            throw new InvalidOperationException("File backup profile đang trống.");
        }

        var packet = JsonSerializer.Deserialize<DataLinkProfileBackupPacket>(json)
            ?? throw new InvalidOperationException("File backup profile không hợp lệ.");

        if (packet.SchemaVersion != 1)
        {
            throw new InvalidOperationException($"Phiên bản backup không hỗ trợ: {packet.SchemaVersion}.");
        }

        var namedData = LoadNamedAll();
        var normalized = packet.Profiles
            .Where(p => !string.IsNullOrWhiteSpace(p.Name))
            .GroupBy(p => p.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => new NamedDataLinkProfile
            {
                Name = g.First().Name.Trim(),
                Profile = g.Last().Profile ?? new DataLinkProfile()
            })
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        namedData[profileKey] = normalized;
        var namedJson = JsonSerializer.Serialize(namedData, JsonOptions);
        File.WriteAllText(_namedProfilePath, namedJson);

        if (!string.IsNullOrWhiteSpace(packet.LastNamedProfile)
            && normalized.Any(p => string.Equals(p.Name, packet.LastNamedProfile, StringComparison.OrdinalIgnoreCase)))
        {
            SaveLastNamedProfileName(profileKey, packet.LastNamedProfile);
        }
        else
        {
            ClearLastNamedProfileName(profileKey);
        }

        if (normalized.Count > 0)
        {
            Save(profileKey, normalized[0].Profile);
        }

        return normalized.Select(p => p.Name).ToList();
    }

    private Dictionary<string, DataLinkProfile> LoadAll()
    {
        if (!File.Exists(_profilePath))
        {
            return new Dictionary<string, DataLinkProfile>();
        }

        var json = File.ReadAllText(_profilePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new Dictionary<string, DataLinkProfile>();
        }

        return JsonSerializer.Deserialize<Dictionary<string, DataLinkProfile>>(json)
            ?? new Dictionary<string, DataLinkProfile>();
    }

    private Dictionary<string, List<NamedDataLinkProfile>> LoadNamedAll()
    {
        if (!File.Exists(_namedProfilePath))
        {
            return new Dictionary<string, List<NamedDataLinkProfile>>();
        }

        var json = File.ReadAllText(_namedProfilePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new Dictionary<string, List<NamedDataLinkProfile>>();
        }

        return JsonSerializer.Deserialize<Dictionary<string, List<NamedDataLinkProfile>>>(json)
            ?? new Dictionary<string, List<NamedDataLinkProfile>>();
    }

    private Dictionary<string, string> LoadLastNamedAll()
    {
        if (!File.Exists(_lastNamedProfilePath))
        {
            return new Dictionary<string, string>();
        }

        var json = File.ReadAllText(_lastNamedProfilePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new Dictionary<string, string>();
        }

        return JsonSerializer.Deserialize<Dictionary<string, string>>(json)
            ?? new Dictionary<string, string>();
    }
}
