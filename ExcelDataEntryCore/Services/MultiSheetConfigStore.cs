using System.IO;
using System.Text.Json;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp.Services;

public sealed class MultiSheetConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _storePath;

    public MultiSheetConfigStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "ExcelDataEntryApp");
        Directory.CreateDirectory(directory);
        _storePath = Path.Combine(directory, "multi-sheet-configs.json");
    }

    public SavedMultiSheetConfig? TryLoad(string workbookPath)
    {
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(_storePath))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(_storePath);
            var data = JsonSerializer.Deserialize<MultiSheetConfigFile>(json);
            if (data?.Configs is null)
            {
                return null;
            }

            var key = NormalizePath(workbookPath);
            return data.Configs.TryGetValue(key, out var config) ? config : null;
        }
        catch
        {
            return null;
        }
    }

    public MultiSheetConfigFile LoadAll()
    {
        if (!File.Exists(_storePath))
        {
            return new MultiSheetConfigFile();
        }

        try
        {
            var json = File.ReadAllText(_storePath);
            return JsonSerializer.Deserialize<MultiSheetConfigFile>(json) ?? new MultiSheetConfigFile();
        }
        catch
        {
            return new MultiSheetConfigFile();
        }
    }

    public void ReplaceAll(MultiSheetConfigFile data)
    {
        var output = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_storePath, output);
    }

    public void Save(string workbookPath, SavedMultiSheetConfig config)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
        {
            return;
        }

        MultiSheetConfigFile data;
        if (File.Exists(_storePath))
        {
            try
            {
                var json = File.ReadAllText(_storePath);
                data = JsonSerializer.Deserialize<MultiSheetConfigFile>(json) ?? new MultiSheetConfigFile();
            }
            catch
            {
                data = new MultiSheetConfigFile();
            }
        }
        else
        {
            data = new MultiSheetConfigFile();
        }

        data.Configs[NormalizePath(workbookPath)] = config;
        var output = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(_storePath, output);
    }

    private static string NormalizePath(string path) =>
        Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
}

public sealed class MultiSheetConfigFile
{
    public Dictionary<string, SavedMultiSheetConfig> Configs { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class SavedMultiSheetConfig
{
    public string NameSheetName { get; set; } = string.Empty;

    public int NameColumnIndex { get; set; } = MultiSheetImportDefaults.NameColumnIndex;

    public bool ShowHiddenSheets { get; set; }

    public bool AutoSkipBlankHeaders { get; set; } = MultiSheetImportDefaults.AutoSkipBlankHeaders;

    public bool SuggestionsDisabled { get; set; } = MultiSheetImportDefaults.SuggestionsDisabled;

    public List<SavedSheetImportConfig> Sheets { get; set; } = [];
}

public sealed class SavedSheetImportConfig
{
    public string SheetName { get; set; } = string.Empty;

    public int HeaderFirstRow { get; set; } = 1;

    public int HeaderLastRow { get; set; } = 3;

    public int FirstDataRow { get; set; } = 4;

    public int SampleRow { get; set; }

    public bool IsNameSheet { get; set; }

    public int NameColumnIndex { get; set; } = MultiSheetImportDefaults.NameColumnIndex;
}
