using System.IO;
using System.Text.Json;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp.Services;

public sealed class AppConfigBackupStore
{
    public const string BackupFileName = "app-config-backup.json";

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string BackupFilePath => Path.Combine(AppContext.BaseDirectory, BackupFileName);

    public bool TryRestoreIfPresent(
        SessionSettingsStore sessionStore,
        MultiSheetConfigStore multiSheetStore,
        HsskExportProfileStore? hsskExportProfileStore = null,
        ImportFileExportProfileStore? importFileExportProfileStore = null)
    {
        if (!File.Exists(BackupFilePath))
        {
            return false;
        }

        try
        {
            var json = File.ReadAllText(BackupFilePath);
            var backup = JsonSerializer.Deserialize<AppConfigBackup>(json);
            if (backup is null)
            {
                return false;
            }

            if (backup.SingleSheet is not null)
            {
                sessionStore.Save(backup.SingleSheet);
            }

            if (backup.MultiSheetConfigs is not null)
            {
                multiSheetStore.ReplaceAll(backup.MultiSheetConfigs);
            }

            if (hsskExportProfileStore is not null && backup.HsskExportProfiles is not null)
            {
                hsskExportProfileStore.ReplaceAll(backup.HsskExportProfiles);
            }

            if (importFileExportProfileStore is not null && backup.ImportFileExportProfiles is not null)
            {
                importFileExportProfileStore.ReplaceAll(backup.ImportFileExportProfiles);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public void SaveBackup(
        SessionSettings singleSheet,
        MultiSheetConfigFile multiSheetConfigs,
        IReadOnlyDictionary<string, HsskExportProfile>? hsskExportProfiles = null,
        IReadOnlyDictionary<string, ImportFileExportProfile>? importFileExportProfiles = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(BackupFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var backup = new AppConfigBackup
            {
                SingleSheet = singleSheet,
                MultiSheetConfigs = multiSheetConfigs,
                HsskExportProfiles = hsskExportProfiles?.ToDictionary(static p => p.Key, static p => p.Value),
                ImportFileExportProfiles = importFileExportProfiles?.ToDictionary(static p => p.Key, static p => p.Value)
            };
            var json = JsonSerializer.Serialize(backup, JsonOptions);
            File.WriteAllText(BackupFilePath, json);
        }
        catch
        {
            // Không chặn thao tác chính nếu không ghi được backup cạnh app.
        }
    }
}

public sealed class AppConfigBackup
{
    public SessionSettings? SingleSheet { get; set; }

    public MultiSheetConfigFile? MultiSheetConfigs { get; set; }

    public Dictionary<string, HsskExportProfile>? HsskExportProfiles { get; set; }

    public Dictionary<string, ImportFileExportProfile>? ImportFileExportProfiles { get; set; }
}
