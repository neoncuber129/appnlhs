using System.IO;
using System.Text.Json;

namespace ExcelDataEntryApp.Services;

public sealed class SessionSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _settingsPath;

    public SessionSettingsStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "ExcelDataEntryApp");
        Directory.CreateDirectory(directory);
        _settingsPath = Path.Combine(directory, "session-settings.json");
    }

    public SessionSettings Load()
    {
        if (!File.Exists(_settingsPath))
        {
            return SessionSettings.Default;
        }

        var json = File.ReadAllText(_settingsPath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return SessionSettings.Default;
        }

        try
        {
            var loaded = JsonSerializer.Deserialize<SessionSettings>(json) ?? SessionSettings.Default;
            loaded.HeaderRowNumber = Math.Max(1, loaded.HeaderRowNumber);
            loaded.NameColumnIndex = Math.Max(1, loaded.NameColumnIndex);
            loaded.SampleRowThreshold = Math.Max(0, loaded.SampleRowThreshold);
            if (string.IsNullOrWhiteSpace(loaded.UiScalePreset))
            {
                loaded.UiScalePreset = SessionSettings.Default.UiScalePreset;
            }

            return loaded;
        }
        catch
        {
            // Self-heal when settings format changes or file is corrupted.
            Save(SessionSettings.Default);
            return SessionSettings.Default;
        }
    }

    public void Save(SessionSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(_settingsPath, json);
    }
}

public sealed class SessionSettings
{
    public int HeaderRowNumber { get; set; } = 3;
    public int NameColumnIndex { get; set; } = 2;
    public int SampleRowThreshold { get; set; } = 4;
    public bool AutoSkipBlankHeaders { get; set; } = true;
    public bool IsAutoSaveEnabled { get; set; } = true;
    public string UiScalePreset { get; set; } = "Vừa";
    /// <summary>Khi true: không tải gợi ý từ cột (giảm tải khi nhập nhiều).</summary>
    public bool SuggestionsDisabled { get; set; }

    /// <summary>Cột giới tính dùng khi xuất HSSK (0 = chưa chọn).</summary>
    public int HsskGenderColumnIndex { get; set; }

    /// <summary>Cột $ không lấy dòng mẫu khi xuất cho Nam (hoặc giới tính trống).</summary>
    public List<int> HsskSkipSampleForMaleColumnIndexes { get; set; } = [];

    public static SessionSettings Default => new();
}
