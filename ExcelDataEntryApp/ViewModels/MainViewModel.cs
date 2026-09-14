using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Data;
using ExcelDataEntryApp.Infrastructure;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;
using Microsoft.Win32;
using OfficeOpenXml;

namespace ExcelDataEntryApp.ViewModels;

public sealed class MainViewModel : ObservableObject, IDisposable
{
    private static readonly Regex LoaiRegex = new(@"\bLo(?:ại|ai)\s*(\d+)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly (string Bg, string Border, string HeaderFg, string InputBg, string InputFg)[] DropdownRowPalette =
    [
        ("#1E3A8A", "#172554", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#065F46", "#022C22", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#7C2D12", "#431407", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#5B21B6", "#3B0764", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#0E7490", "#164E63", "#FFFFFF", "#F8FAFC", "#0F172A"),
        ("#854D0E", "#422006", "#FFFFFF", "#F8FAFC", "#0F172A"),
    ];

    private readonly ExcelWorkbookService _workbookService = new();
    private readonly HeaderProfileStore _profileStore = new();
    private readonly DataLinkProfileStore _dataLinkProfileStore = new();
    private readonly ExcelDataLinkService _dataLinkService = new();
    private readonly SessionSettingsStore _sessionSettingsStore = new();
    private readonly MultiSheetConfigStore _multiSheetConfigStore = new();
    private readonly AppConfigBackupStore _appConfigBackupStore = new();
    private readonly HsskExportProfileStore _hsskExportProfileStore = new();
    private readonly ImportFileExportProfileStore _importFileExportProfileStore = new();
    private readonly RelayCommand _selectSheetCommand;
    private readonly RelayCommand _applyHeaderRowCommand;
    private readonly RelayCommand _loadRecordsCommand;
    private readonly RelayCommand _toggleAllHeadersCommand;
    private readonly RelayCommand _addRecordCommand;
    private readonly RelayCommand _saveCommand;
    private readonly RelayCommand _saveAsCommand;
    private readonly RelayCommand<RecordItem> _editRecordCommand;
    private readonly RelayCommand<RecordItem> _deleteRecordCommand;

    private readonly Dictionary<CellKey, string> _pendingCellValues = [];
    private MultiSheetImportSession? _multiSheetSession;
    private bool _isMultiSheetMode;
    private bool _suppressSheetSelectionChange;
    private HashSet<string> _hiddenHeadersBeforeShowAll = [];
    private int _pendingRowIndex = -1;
    private bool _suppressAutosave;
    private bool _isShowAllTemporarilyEnabled;
    private bool _isBulkUpdatingHeaderVisibility;
    private bool _isWorkbookLoaded;
    private bool _isRealtimeRefreshInProgress;
    private string? _profileKey;

    private string _statusMessage = "Chọn file Excel để bắt đầu.";
    private string _filePath = string.Empty;
    private string _searchText = string.Empty;
    private string _selectedSheet = string.Empty;
    private int _headerRowNumber = 3;
    private int _nameColumnIndex = 2;
    private int _sampleRowThreshold = 4;
    private bool _autoSkipBlankHeaders = true;
    private bool _isAutoSaveEnabled = true;
    private string _uiScalePreset = "Vừa";
    private bool _suggestionsDisabled;
    private RecordItem? _selectedRecord;

    public MainViewModel()
    {
        _appConfigBackupStore.TryRestoreIfPresent(
            _sessionSettingsStore,
            _multiSheetConfigStore,
            _hsskExportProfileStore,
            _importFileExportProfileStore);
        var lastSession = _sessionSettingsStore.Load();
        _headerRowNumber = Math.Max(1, lastSession.HeaderRowNumber);
        _nameColumnIndex = Math.Max(1, lastSession.NameColumnIndex);
        _sampleRowThreshold = Math.Max(0, lastSession.SampleRowThreshold);
        _autoSkipBlankHeaders = lastSession.AutoSkipBlankHeaders;
        _isAutoSaveEnabled = lastSession.IsAutoSaveEnabled;
        _uiScalePreset = NormalizeUiScalePreset(lastSession.UiScalePreset);
        _suggestionsDisabled = lastSession.SuggestionsDisabled;
        HeaderColumns = [];
        Sheets = [];
        Records = [];
        EditableFields = [];

        UploadFileCommand = new RelayCommand(UploadFile);
        _selectSheetCommand = new RelayCommand(SelectSheet, () => !string.IsNullOrWhiteSpace(SelectedSheet));
        _applyHeaderRowCommand = new RelayCommand(ApplyHeaderRow, () => HeaderRowNumber > 0);
        _loadRecordsCommand = new RelayCommand(LoadRecords, () => GetActiveNameColumnIndex() > 0 && HeaderColumns.Count > 0);
        _toggleAllHeadersCommand = new RelayCommand(ToggleAllHeaders, () => HeaderColumns.Count > 0);
        _addRecordCommand = new RelayCommand(AddRecord, () => _isWorkbookLoaded && HeaderColumns.Count > 0 && GetActiveNameColumnIndex() > 0);
        _saveCommand = new RelayCommand(SaveNow, () => _isWorkbookLoaded);
        _saveAsCommand = new RelayCommand(SaveAs, () => _isWorkbookLoaded);
        _editRecordCommand = new RelayCommand<RecordItem>(EditRecord, r => _isWorkbookLoaded && (r is not null || SelectedRecord is not null) && GetActiveNameColumnIndex() > 0);
        _deleteRecordCommand = new RelayCommand<RecordItem>(DeleteRecord, r => _isWorkbookLoaded && (r is not null || SelectedRecord is not null) && GetActiveNameColumnIndex() > 0);

        SelectSheetCommand = _selectSheetCommand;
        ApplyHeaderRowCommand = _applyHeaderRowCommand;
        LoadRecordsCommand = _loadRecordsCommand;
        ToggleAllHeadersCommand = _toggleAllHeadersCommand;
        AddRecordCommand = _addRecordCommand;
        SaveCommand = _saveCommand;
        SaveAsCommand = _saveAsCommand;
        EditRecordCommand = _editRecordCommand;
        DeleteRecordCommand = _deleteRecordCommand;

        RecordsView = CollectionViewSource.GetDefaultView(Records);
        RecordsView.Filter = FilterRecords;
    }

    public RelayCommand UploadFileCommand { get; }
    public RelayCommand SelectSheetCommand { get; }
    public RelayCommand ApplyHeaderRowCommand { get; }
    public RelayCommand LoadRecordsCommand { get; }
    public RelayCommand ToggleAllHeadersCommand { get; }
    public RelayCommand AddRecordCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand SaveAsCommand { get; }
    public RelayCommand<RecordItem> EditRecordCommand { get; }
    public RelayCommand<RecordItem> DeleteRecordCommand { get; }

    public ObservableCollection<string> Sheets { get; }
    public ObservableCollection<HeaderDefinition> HeaderColumns { get; }
    public ObservableCollection<RecordItem> Records { get; }
    public ObservableCollection<EditableField> EditableFields { get; }
    public ICollectionView RecordsView { get; }

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public string FilePath
    {
        get => _filePath;
        set => SetProperty(ref _filePath, value);
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetProperty(ref _searchText, value))
            {
                RecordsView.Refresh();
            }
        }
    }

    public string SelectedSheet
    {
        get => _selectedSheet;
        set
        {
            if (SetProperty(ref _selectedSheet, value))
            {
                _selectSheetCommand.RaiseCanExecuteChanged();
                if (_suppressSheetSelectionChange || IsMultiSheetMode)
                {
                    return;
                }

                if (_isWorkbookLoaded && !_isRealtimeRefreshInProgress && !string.IsNullOrWhiteSpace(_selectedSheet))
                {
                    SelectSheet();
                }
            }
        }
    }

    public int HeaderRowNumber
    {
        get => _headerRowNumber;
        set
        {
            if (SetProperty(ref _headerRowNumber, value))
            {
                _applyHeaderRowCommand.RaiseCanExecuteChanged();
                RefreshOnHeaderDefinitionChanged();
            }
        }
    }

    public int NameColumnIndex
    {
        get => _nameColumnIndex;
        set
        {
            if (SetProperty(ref _nameColumnIndex, value))
            {
                _loadRecordsCommand.RaiseCanExecuteChanged();
                _editRecordCommand.RaiseCanExecuteChanged();
                _deleteRecordCommand.RaiseCanExecuteChanged();
                RefreshOnNameColumnChanged();
            }
        }
    }

    public int SampleRowThreshold
    {
        get => _sampleRowThreshold;
        set
        {
            if (SetProperty(ref _sampleRowThreshold, value) && SelectedRecord is not null)
            {
                _workbookService.InvalidateActiveSheetSuggestionCaches();
                LoadEditableFieldsForSelectedRecord();
            }
        }
    }

    public bool AutoSkipBlankHeaders
    {
        get => _autoSkipBlankHeaders;
        set
        {
            if (SetProperty(ref _autoSkipBlankHeaders, value))
            {
                RefreshOnHeaderDefinitionChanged();
            }
        }
    }

    public bool IsAutoSaveEnabled
    {
        get => _isAutoSaveEnabled;
        set
        {
            if (SetProperty(ref _isAutoSaveEnabled, value))
            {
                if (_isAutoSaveEnabled)
                {
                    FlushPendingChanges(saveToDisk: true);
                    StatusMessage = "Đã bật auto lưu.";
                }
                else
                {
                    StatusMessage = "Đã tắt auto lưu. Hãy bấm Lưu để ghi file.";
                }
            }
        }
    }

    public bool IsSuggestionsEnabled
    {
        get => !_suggestionsDisabled;
        set
        {
            if (!SetProperty(ref _suggestionsDisabled, !value))
            {
                return;
            }

            if (_suggestionsDisabled)
            {
                _workbookService.InvalidateActiveSheetSuggestionCaches();
            }

            if (SelectedRecord is not null)
            {
                LoadEditableFieldsForSelectedRecord();
            }
        }
    }

    public IReadOnlyList<string> UiScalePresets { get; } = ["Nhỏ", "Vừa", "To", "Rất to"];

    public string UiScalePreset
    {
        get => _uiScalePreset;
        set
        {
            var normalized = NormalizeUiScalePreset(value);
            if (SetProperty(ref _uiScalePreset, normalized))
            {
                OnPropertyChanged(nameof(UiScaleFactor));
            }
        }
    }

    public double UiScaleFactor => UiScalePreset switch
    {
        "Nhỏ" => 0.9,
        "To" => 1.12,
        "Rất to" => 1.25,
        _ => 1.0
    };

    public bool AreAllHeadersVisible => HeaderColumns.Count > 0 && HeaderColumns.All(h => h.IsVisible);

    public string ToggleAllHeadersLabel => _isShowAllTemporarilyEnabled ? "Tắt hiện tất cả" : "Hiện tất cả";

    public bool IsMultiSheetMode
    {
        get => _isMultiSheetMode;
        private set
        {
            if (SetProperty(ref _isMultiSheetMode, value))
            {
                OnPropertyChanged(nameof(IsSingleSheetMode));
                OnPropertyChanged(nameof(IsSingleSheetToolbarEnabled));
            }
        }
    }

    public bool IsSingleSheetMode => !IsMultiSheetMode;

    public bool IsSingleSheetToolbarEnabled => !IsMultiSheetMode;

    public bool IsWorkbookLoaded => _isWorkbookLoaded;

    public string ConfigBackupFilePath => _appConfigBackupStore.BackupFilePath;

    public string MultiSheetModeCaption =>
        IsMultiSheetMode && _multiSheetSession is not null
            ? $"Đang ghép {_multiSheetSession.Sheets.Count} sheet (tên: {_multiSheetSession.NameSheetName})"
            : string.Empty;

    /// <summary>Chỉ báo phía trên form: đang nhập cho ai (cột tên / danh sách bên trái).</summary>
    public string CurrentRecordEntryCaption
    {
        get
        {
            if (SelectedRecord is null)
            {
                return "Chưa chọn dòng — chọn một tên trong danh sách bên trái.";
            }

            return $"Đang nhập cho \"{SelectedRecord.KeyDisplay}\"";
        }
    }

    public RecordItem? SelectedRecord
    {
        get => _selectedRecord;
        set
        {
            if (!IsSameRecord(_selectedRecord, value))
            {
                FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
                _pendingRowIndex = -1;
            }

            if (SetProperty(ref _selectedRecord, value))
            {
                LoadEditableFieldsForSelectedRecord();
            }
        }
    }

    public IReadOnlyList<WorksheetInfo> GetWorksheetInfos(bool includeHidden) =>
        _workbookService.GetWorksheetInfos(includeHidden);

    public IReadOnlyList<HeaderCell> ReadHeaderBlock(string sheetName, int firstRow, int lastRow) =>
        _workbookService.ReadHeaderBlock(sheetName, firstRow, lastRow);

    public void ApplyMultiSheetImport(MultiSheetImportSession session)
    {
        if (!_isWorkbookLoaded || session.Sheets.Count == 0)
        {
            return;
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged -= HeaderOnPropertyChanged;
        }

        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;

        _multiSheetSession = session;

        _suppressSheetSelectionChange = true;
        try
        {
            SelectedSheet = session.NameSheetName;
        }
        finally
        {
            _suppressSheetSelectionChange = false;
        }

        _workbookService.SelectWorksheet(session.NameSheetName);
        IsMultiSheetMode = true;
        OnPropertyChanged(nameof(MultiSheetModeCaption));
        BuildMergedHeaderColumns();
        LoadRecordsMultiSheet();
        SaveMultiSheetConfig(session);
        StatusMessage = $"Đang nhập {session.Sheets.Count} sheet ghép. Danh sách tên theo sheet '{session.NameSheetName}'.";
    }

    public SavedMultiSheetConfig? TryGetSavedMultiSheetConfig()
    {
        if (!_isWorkbookLoaded || string.IsNullOrWhiteSpace(_workbookService.WorkbookPath))
        {
            return null;
        }

        return _multiSheetConfigStore.TryLoad(_workbookService.WorkbookPath);
    }

    private void SaveMultiSheetConfig(MultiSheetImportSession session)
    {
        if (string.IsNullOrWhiteSpace(_workbookService.WorkbookPath))
        {
            return;
        }

        _multiSheetConfigStore.Save(_workbookService.WorkbookPath, new SavedMultiSheetConfig
        {
            NameSheetName = session.NameSheetName,
            NameColumnIndex = session.NameColumnIndex,
            ShowHiddenSheets = session.ShowHiddenSheets,
            AutoSkipBlankHeaders = session.AutoSkipBlankHeaders,
            SuggestionsDisabled = session.SuggestionsDisabled,
            Sheets = session.Sheets.Select(s => new SavedSheetImportConfig
            {
                SheetName = s.SheetName,
                HeaderFirstRow = s.HeaderFirstRow,
                HeaderLastRow = s.HeaderLastRow,
                FirstDataRow = s.FirstDataRow,
                SampleRow = s.SampleRow,
                IsNameSheet = string.Equals(s.SheetName, session.NameSheetName, StringComparison.Ordinal),
                NameColumnIndex = string.Equals(s.SheetName, session.NameSheetName, StringComparison.Ordinal)
                    ? session.NameColumnIndex
                    : s.NameColumnIndex
            }).ToList()
        });

        PersistAppConfigBackup();
    }

    private void TryRestoreSavedMultiSheetConfig()
    {
        var saved = TryGetSavedMultiSheetConfig();
        if (saved is null || saved.Sheets.Count < 2)
        {
            return;
        }

        var workbookSheets = new HashSet<string>(_workbookService.GetWorksheetNames(), StringComparer.Ordinal);
        if (saved.Sheets.Any(s => !workbookSheets.Contains(s.SheetName)))
        {
            return;
        }

        var session = new MultiSheetImportSession
        {
            NameSheetName = saved.NameSheetName,
            NameColumnIndex = saved.NameColumnIndex,
            ShowHiddenSheets = saved.ShowHiddenSheets,
            AutoSkipBlankHeaders = saved.AutoSkipBlankHeaders,
            SuggestionsDisabled = saved.SuggestionsDisabled,
            Sheets = saved.Sheets.Select(s => new SheetImportConfig
            {
                SheetName = s.SheetName,
                HeaderFirstRow = s.HeaderFirstRow,
                HeaderLastRow = s.HeaderLastRow,
                FirstDataRow = s.FirstDataRow,
                SampleRow = s.SampleRow,
                IsNameSheet = s.IsNameSheet,
                NameColumnIndex = s.NameColumnIndex
            }).ToList()
        };

        ApplyMultiSheetImport(session);
    }

    public void ApplySingleSheetConfig(
        int headerRow,
        int nameColumn,
        int sampleRow,
        bool autoSkipBlankHeaders,
        bool suggestionsEnabled,
        bool autoSaveEnabled)
    {
        HeaderRowNumber = headerRow;
        NameColumnIndex = nameColumn;
        SampleRowThreshold = sampleRow;
        AutoSkipBlankHeaders = autoSkipBlankHeaders;
        IsSuggestionsEnabled = suggestionsEnabled;
        IsAutoSaveEnabled = autoSaveEnabled;

        if (!IsMultiSheetMode && _isWorkbookLoaded && !string.IsNullOrWhiteSpace(SelectedSheet))
        {
            ApplyHeaderRow();
            LoadRecords();
        }

        PersistSingleSheetSettingsAndBackup();
    }

    public void ExportConfigBackup()
    {
        if (!IsMultiSheetMode)
        {
            PersistSingleSheetSettingsAndBackup();
            return;
        }

        PersistAppConfigBackup();
    }

    public int GetSampleRowForSheet(string sheetName)
    {
        if (IsMultiSheetMode && _multiSheetSession?.GetSheetConfig(sheetName) is { } cfg)
        {
            return ResolveSampleRow(cfg.SampleRow, cfg.FirstDataRow, cfg.HeaderLastRow);
        }

        return SampleRowThreshold;
    }

    private SessionSettings BuildSingleSheetSessionSettings() => new()
    {
        HeaderRowNumber = HeaderRowNumber,
        NameColumnIndex = NameColumnIndex,
        SampleRowThreshold = SampleRowThreshold,
        AutoSkipBlankHeaders = AutoSkipBlankHeaders,
        IsAutoSaveEnabled = IsAutoSaveEnabled,
        UiScalePreset = UiScalePreset,
        SuggestionsDisabled = _suggestionsDisabled
    };

    public HsskExportProfile GetHsskExportProfile()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return new HsskExportProfile();
        }

        var profile = _hsskExportProfileStore.TryLoad(_profileKey);
        if (profile is not null)
        {
            profile.SkipSampleForMaleColumnIndexes ??= [];
            return profile;
        }

        var session = _sessionSettingsStore.Load();
        return new HsskExportProfile
        {
            GenderColumnIndex = Math.Max(0, session.HsskGenderColumnIndex),
            SkipSampleForMaleColumnIndexes = session.HsskSkipSampleForMaleColumnIndexes ?? []
        };
    }

    public void SaveHsskExportProfile(HsskExportProfile profile)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        profile.SkipSampleForMaleColumnIndexes = profile.SkipSampleForMaleColumnIndexes
            .Where(i => i > 0)
            .Distinct()
            .OrderBy(i => i)
            .ToList();
        profile.GenderColumnIndex = Math.Max(0, profile.GenderColumnIndex);
        _hsskExportProfileStore.Save(_profileKey, profile);
        PersistAppConfigBackup();
    }

    public bool SaveHsskSkipSampleForMaleColumns(IEnumerable<int> columnIndexes)
    {
        var profile = GetHsskExportProfile();
        profile.SkipSampleForMaleColumnIndexes = columnIndexes
            .Where(i => i > 0)
            .Distinct()
            .OrderBy(i => i)
            .ToList();
        SaveHsskExportProfile(profile);
        return !string.IsNullOrWhiteSpace(_profileKey);
    }

    public int HsskGenderColumnIndex
    {
        get => GetHsskExportProfile().GenderColumnIndex;
        set
        {
            var profile = GetHsskExportProfile();
            profile.GenderColumnIndex = Math.Max(0, value);
            SaveHsskExportProfile(profile);
        }
    }

    public IReadOnlyList<HsskFemaleColumnOption> GetHsskColumnOptionsForExport()
    {
        var savedColumns = GetHsskExportProfile().SkipSampleForMaleColumnIndexes.ToHashSet();
        return HeaderColumns
            .GroupBy(h => h.ColumnIndex)
            .OrderBy(g => g.Key)
            .Select(g => new HsskFemaleColumnOption
            {
                ColumnIndex = g.Key,
                HeaderName = g.First().Name,
                IsSelected = savedColumns.Contains(g.Key)
            })
            .ToList();
    }

    public IReadOnlyList<int> ScanMarkerColumnIndexes() => ScanMarkerColumnIndexesOnSheet(_workbookService.ActiveSheetName ?? SelectedSheet);

    public IReadOnlyList<int> ScanHsskFemaleOnlyColumnIndexes() => ScanMarkerColumnIndexes();

    public ImportFileExportProfile GetImportFileExportProfile()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return new ImportFileExportProfile();
        }

        var profile = _importFileExportProfileStore.TryLoad(_profileKey);
        if (profile is not null)
        {
            profile.SkipSampleColumnIndexes ??= [];
            return profile;
        }

        var hsskProfile = GetHsskExportProfile();
        return new ImportFileExportProfile
        {
            GenderColumnIndex = hsskProfile.GenderColumnIndex,
            SkipSampleColumnIndexes = []
        };
    }

    public void SaveImportFileExportProfile(ImportFileExportProfile profile)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        profile.GenderColumnIndex = Math.Max(0, profile.GenderColumnIndex);
        profile.SkipSampleColumnIndexes = profile.SkipSampleColumnIndexes
            .Where(i => i > 0)
            .Distinct()
            .OrderBy(i => i)
            .ToList();
        _importFileExportProfileStore.Save(_profileKey, profile);
        PersistAppConfigBackup();
    }

    public int ImportFileGenderColumnIndex
    {
        get => GetImportFileExportProfile().GenderColumnIndex;
        set
        {
            var profile = GetImportFileExportProfile();
            profile.GenderColumnIndex = Math.Max(0, value);
            SaveImportFileExportProfile(profile);
        }
    }

    public bool SaveImportFileSkipSampleColumns(IEnumerable<int> columnIndexes)
    {
        var profile = GetImportFileExportProfile();
        profile.SkipSampleColumnIndexes = columnIndexes
            .Where(i => i > 0)
            .Distinct()
            .OrderBy(i => i)
            .ToList();
        SaveImportFileExportProfile(profile);
        return !string.IsNullOrWhiteSpace(_profileKey);
    }

    public bool HasExportProfileKey => !string.IsNullOrWhiteSpace(_profileKey);

    public IReadOnlyList<HsskFemaleColumnOption> GetImportFileColumnOptionsForExport()
    {
        var savedColumns = GetImportFileExportProfile().SkipSampleColumnIndexes.ToHashSet();
        return HeaderColumns
            .GroupBy(h => h.ColumnIndex)
            .OrderBy(g => g.Key)
            .Select(g => new HsskFemaleColumnOption
            {
                ColumnIndex = g.Key,
                HeaderName = g.First().Name,
                IsSelected = savedColumns.Contains(g.Key)
            })
            .ToList();
    }

    private IReadOnlyList<int> ScanMarkerColumnIndexesOnSheet(string? sheetName)
    {
        if (!_isWorkbookLoaded || string.IsNullOrWhiteSpace(sheetName))
        {
            return [];
        }

        return _workbookService.FindColumnsWithMarkerInRows(
            sheetName,
            1,
            2,
            HsskExportGenderHelper.FemaleColumnMarker);
    }

    public HeaderDefinition? TryGuessHsskGenderHeader() =>
        HeaderColumns.FirstOrDefault(h =>
            h.Name.Contains("giới", StringComparison.OrdinalIgnoreCase)
            || h.Name.Contains("gioi", StringComparison.OrdinalIgnoreCase)
            || h.Name.Contains("giới tính", StringComparison.OrdinalIgnoreCase));

    private void PersistSingleSheetSettingsAndBackup()
    {
        _sessionSettingsStore.Save(BuildSingleSheetSessionSettings());
        PersistAppConfigBackup();
    }

    private void PersistAppConfigBackup()
    {
        _appConfigBackupStore.SaveBackup(
            _sessionSettingsStore.Load(),
            _multiSheetConfigStore.LoadAll(),
            _hsskExportProfileStore.LoadAll(),
            _importFileExportProfileStore.LoadAll());
    }

    public string HsskExportProfilesFilePath => _hsskExportProfileStore.ProfileFilePath;

    public string ImportFileExportProfilesFilePath => _importFileExportProfileStore.ProfileFilePath;

    public string AppConfigBackupFilePath => _appConfigBackupStore.BackupFilePath;

    public bool PersistHsskExportProfileToAppFolder()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return false;
        }

        SaveHsskExportProfile(GetHsskExportProfile());
        PersistAppConfigBackup();
        return true;
    }

    public bool PersistImportFileExportProfileToAppFolder()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return false;
        }

        SaveImportFileExportProfile(GetImportFileExportProfile());
        PersistAppConfigBackup();
        return true;
    }

    private int GetActiveNameColumnIndex() =>
        IsMultiSheetMode && _multiSheetSession is not null
            ? _multiSheetSession.NameColumnIndex
            : NameColumnIndex;

    private bool AreSuggestionsEnabledForCurrentMode() =>
        IsMultiSheetMode && _multiSheetSession is not null
            ? !_multiSheetSession.SuggestionsDisabled
            : !_suggestionsDisabled;

    private static int ResolveSampleRow(int configuredSampleRow, int firstDataRow, int headerLastRow)
    {
        if (configuredSampleRow > 0)
        {
            return configuredSampleRow;
        }

        if (firstDataRow > headerLastRow)
        {
            return firstDataRow;
        }

        return Math.Max(1, headerLastRow);
    }

    private int ResolveSingleSheetFirstDataRow()
    {
        var afterHeader = HeaderRowNumber + 1;
        return SampleRowThreshold > 0
            ? Math.Max(afterHeader, SampleRowThreshold)
            : afterHeader;
    }

    private int ResolveSingleSheetSortStartRow() => Math.Max(1, HeaderRowNumber + 1);

    private int ResolveSingleSheetSortExcludedRow() => SampleRowThreshold > 0 ? SampleRowThreshold : 0;

    public void ExitMultiSheetMode()
    {
        if (!IsMultiSheetMode)
        {
            return;
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        ClearMultiSheetState();
        IsMultiSheetMode = false;
        OnPropertyChanged(nameof(MultiSheetModeCaption));

        if (_isWorkbookLoaded && !string.IsNullOrWhiteSpace(SelectedSheet))
        {
            SelectSheet();
        }
    }

    private void ClearMultiSheetState()
    {
        _multiSheetSession = null;
        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged -= HeaderOnPropertyChanged;
        }

        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
    }

    private static bool IsSameRecord(RecordItem? left, RecordItem? right)
    {
        if (left is null && right is null)
        {
            return true;
        }

        if (left is null || right is null)
        {
            return false;
        }

        if (left.LogicalIndex >= 0 && right.LogicalIndex >= 0)
        {
            return left.LogicalIndex == right.LogicalIndex;
        }

        return left.RowIndex == right.RowIndex;
    }

    private void UploadFile()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn file Excel",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        _workbookService.OpenWorkbook(dialog.FileName);
        _isWorkbookLoaded = true;
        OnPropertyChanged(nameof(IsWorkbookLoaded));
        IsMultiSheetMode = false;
        _multiSheetSession = null;
        OnPropertyChanged(nameof(MultiSheetModeCaption));
        FilePath = dialog.FileName;
        _saveCommand.RaiseCanExecuteChanged();
        _saveAsCommand.RaiseCanExecuteChanged();
        _addRecordCommand.RaiseCanExecuteChanged();
        _editRecordCommand.RaiseCanExecuteChanged();
        _deleteRecordCommand.RaiseCanExecuteChanged();

        Sheets.Clear();
        foreach (var sheet in _workbookService.GetWorksheetNames())
        {
            Sheets.Add(sheet);
        }

        SelectedSheet = Sheets.FirstOrDefault() ?? string.Empty;
        StatusMessage = $"Đã mở file: {Path.GetFileName(dialog.FileName)}.";
        TryRestoreSavedMultiSheetConfig();
    }

    private void SelectSheet()
    {
        if (string.IsNullOrWhiteSpace(SelectedSheet))
        {
            return;
        }

        _workbookService.SelectWorksheet(SelectedSheet);
        _isRealtimeRefreshInProgress = true;
        _isRealtimeRefreshInProgress = false;
        _workbookService.InvalidateActiveSheetCaches();
        HeaderColumns.Clear();
        Records.Clear();
        EditableFields.Clear();
        SelectedRecord = null;
        ApplyHeaderRow();
        LoadRecords();
        StatusMessage = $"Đã chọn sheet {SelectedSheet}. Dữ liệu đã cập nhật realtime.";
    }

    private void ApplyHeaderRow()
    {
        if (!_isWorkbookLoaded || string.IsNullOrWhiteSpace(SelectedSheet) || HeaderRowNumber <= 0)
        {
            return;
        }

        var headers = _workbookService.ReadHeaderRow(HeaderRowNumber);
        if (AutoSkipBlankHeaders)
        {
            headers = headers.Where(h => !string.Equals(h.Name.Trim(), "blank", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged -= HeaderOnPropertyChanged;
        }

        HeaderColumns.Clear();
        foreach (var item in headers)
        {
            HeaderColumns.Add(new HeaderDefinition { ColumnIndex = item.ColumnIndex, Name = item.Name, IsVisible = true });
        }

        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged += HeaderOnPropertyChanged;
        }

        _workbookService.InvalidateActiveSheetCaches();
        _profileKey = _profileStore.BuildKey(_workbookService.WorkbookPath, SelectedSheet, HeaderColumns.Select(h => h.Name));
        ApplyProfileIfAny();

        StatusMessage = $"Đã nhận diện {HeaderColumns.Count} header.";
        _loadRecordsCommand.RaiseCanExecuteChanged();
        _toggleAllHeadersCommand.RaiseCanExecuteChanged();
        _addRecordCommand.RaiseCanExecuteChanged();
        _editRecordCommand.RaiseCanExecuteChanged();
        _deleteRecordCommand.RaiseCanExecuteChanged();
        OnPropertyChanged(nameof(AreAllHeadersVisible));
        OnPropertyChanged(nameof(ToggleAllHeadersLabel));
    }

    private void ReloadRecordItemsFromWorkbook()
    {
        if (IsMultiSheetMode)
        {
            LoadRecordsMultiSheet(preserveSelection: true);
            return;
        }

        if (!_isWorkbookLoaded || string.IsNullOrWhiteSpace(SelectedSheet) || NameColumnIndex <= 0 || HeaderColumns.Count == 0)
        {
            return;
        }

        var rows = _workbookService.GetRecordRows(HeaderRowNumber, NameColumnIndex, HeaderColumns.Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
        Records.Clear();
        foreach (var row in rows)
        {
            Records.Add(new RecordItem
            {
                RowIndex = row.RowIndex,
                KeyDisplay = row.KeyDisplay,
                TooltipPreview = row.TooltipPreview
            });
        }

        RecordsView.Refresh();
        StatusMessage = $"Đã tải {Records.Count} dòng. Dữ liệu bên phải cập nhật theo dòng chọn.";
    }

    private void BuildMergedHeaderColumns()
    {
        if (_multiSheetSession is null)
        {
            return;
        }

        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged -= HeaderOnPropertyChanged;
        }

        HeaderColumns.Clear();

        foreach (var sheetCfg in _multiSheetSession.Sheets)
        {
            var headers = _workbookService.ReadHeaderBlock(
                sheetCfg.SheetName,
                sheetCfg.HeaderFirstRow,
                sheetCfg.HeaderLastRow);

            foreach (var headerCell in headers)
            {
                if (_multiSheetSession.AutoSkipBlankHeaders
                    && string.Equals(headerCell.Name.Trim(), "blank", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                HeaderColumns.Add(new HeaderDefinition
                {
                    SheetName = sheetCfg.SheetName,
                    ColumnIndex = headerCell.ColumnIndex,
                    Name = headerCell.Name,
                    HeaderFirstRow = sheetCfg.HeaderFirstRow,
                    HeaderLastRow = sheetCfg.HeaderLastRow,
                    IsVisible = true
                });
            }
        }

        foreach (var header in HeaderColumns)
        {
            header.PropertyChanged += HeaderOnPropertyChanged;
        }

        _profileKey = _profileStore.BuildKey(
            _workbookService.WorkbookPath,
            "multi:" + string.Join("|", _multiSheetSession.Sheets.Select(s => s.SheetName)),
            HeaderColumns.Select(h => $"{h.SheetName}:{h.Name}"));
        ApplyProfileIfAny();
        _loadRecordsCommand.RaiseCanExecuteChanged();
        _toggleAllHeadersCommand.RaiseCanExecuteChanged();
        _addRecordCommand.RaiseCanExecuteChanged();
        _editRecordCommand.RaiseCanExecuteChanged();
        _deleteRecordCommand.RaiseCanExecuteChanged();
        OnPropertyChanged(nameof(AreAllHeadersVisible));
        OnPropertyChanged(nameof(ToggleAllHeadersLabel));
    }

    private void LoadRecordsMultiSheet(bool preserveSelection = false)
    {
        if (_multiSheetSession is null || GetActiveNameColumnIndex() <= 0)
        {
            return;
        }

        var nameSheet = _multiSheetSession.NameSheet;
        if (nameSheet is null)
        {
            return;
        }

        var nameColumnIndex = GetActiveNameColumnIndex();
        var nameHeaders = HeaderColumns
            .Where(h => string.Equals(h.SheetName, nameSheet.SheetName, StringComparison.Ordinal))
            .Select(h => new HeaderCell(h.ColumnIndex, h.Name))
            .ToList();

        var previousLogical = preserveSelection ? SelectedRecord?.LogicalIndex : null;
        var rows = _workbookService.GetLogicalRecordRowsFromSheet(
            nameSheet.SheetName,
            nameSheet.FirstDataRow,
            nameColumnIndex,
            nameHeaders);

        Records.Clear();
        foreach (var row in rows)
        {
            Records.Add(new RecordItem
            {
                LogicalIndex = row.LogicalIndex,
                RowIndex = row.RowIndex,
                KeyDisplay = row.KeyDisplay,
                TooltipPreview = row.TooltipPreview
            });
        }

        RecordsView.Refresh();
        SelectedRecord = previousLogical is int idx
            ? Records.FirstOrDefault(r => r.LogicalIndex == idx) ?? Records.FirstOrDefault()
            : Records.FirstOrDefault();
        StatusMessage = $"Đã tải {Records.Count} dòng từ sheet tên '{nameSheet.SheetName}'.";
    }

    private void LoadRecords()
    {
        if (IsMultiSheetMode)
        {
            LoadRecordsMultiSheet();
            return;
        }

        ReloadRecordItemsFromWorkbook();
        SelectedRecord = Records.FirstOrDefault();
    }

    private void LoadEditableFieldsForSelectedRecord()
    {
        EditableFields.Clear();
        if (SelectedRecord is null)
        {
            OnPropertyChanged(nameof(CurrentRecordEntryCaption));
            return;
        }

        var visibleHeaders = HeaderColumns.Where(h => h.IsVisible).ToList();
        var singleSheetValues = IsMultiSheetMode
            ? null
            : _workbookService.ReadRowValues(SelectedRecord.RowIndex, visibleHeaders.Select(h => h.ColumnIndex));
        _suppressAutosave = true;
        try
        {
            var fieldDrafts = new List<(HeaderDefinition Header, int SourceRow, string SheetName, string CellValue, IReadOnlyList<string> DropdownOptions, IReadOnlyList<int> ParentDropdownColumns, IReadOnlyList<string> SuggestionOptions)>();
            var rowValueOverrides = new Dictionary<CellKey, string>();
            var dropdownColorSlot = 0;
            var dropdownColorGroups = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var header in visibleHeaders)
            {
                var sourceRow = IsMultiSheetMode && _multiSheetSession is not null
                    ? _multiSheetSession.GetDataRowForLogicalIndex(header.SheetName, SelectedRecord.LogicalIndex)
                    : SelectedRecord.RowIndex;
                var sheetName = IsMultiSheetMode ? header.SheetName : (_workbookService.ActiveSheetName ?? string.Empty);
                var values = IsMultiSheetMode
                    ? _workbookService.ReadRowValues(sheetName, sourceRow, [header.ColumnIndex])
                    : singleSheetValues!;
                var cellValue = values.TryGetValue(header.ColumnIndex, out var value) ? value : string.Empty;
                rowValueOverrides[new CellKey(sheetName, sourceRow, header.ColumnIndex)] = cellValue;
            }

            foreach (var header in visibleHeaders)
            {
                var sourceRow = IsMultiSheetMode && _multiSheetSession is not null
                    ? _multiSheetSession.GetDataRowForLogicalIndex(header.SheetName, SelectedRecord.LogicalIndex)
                    : SelectedRecord.RowIndex;
                var sheetName = IsMultiSheetMode ? header.SheetName : (_workbookService.ActiveSheetName ?? string.Empty);
                var dependency = _workbookService.GetDropdownDependencyInfo(sheetName, sourceRow, header.ColumnIndex);
                var dropdownOptions = dependency.IsDependent
                    ? _workbookService.GetDropdownOptions(sheetName, sourceRow, header.ColumnIndex, rowValueOverrides)
                    : IsMultiSheetMode
                        ? _workbookService.GetDropdownOptions(sheetName, sourceRow, header.ColumnIndex)
                        : _workbookService.GetDropdownOptions(SelectedRecord.RowIndex, header.ColumnIndex);
                dropdownOptions = NormalizeDropdownOptions(dropdownOptions);
                if (dropdownOptions.Count > 0)
                {
                    var colorGroupKey = BuildDropdownColorGroupKey(sheetName, header.Name);
                    if (!dropdownColorGroups.ContainsKey(colorGroupKey))
                    {
                        dropdownColorGroups[colorGroupKey] = dropdownColorSlot++;
                    }
                }

                var sampleRow = GetSampleRowForSheet(header.SheetName);
                var suggestionOptions = AreSuggestionsEnabledForCurrentMode()
                    && dropdownOptions.Count == 0
                    && !dependency.IsDependent
                    && (IsMultiSheetMode
                        ? _workbookService.ColumnHasSuggestionData(header.SheetName, header.ColumnIndex, sampleRow)
                        : _workbookService.ColumnHasSuggestionData(header.ColumnIndex, sampleRow))
                    ? IsMultiSheetMode
                        ? _workbookService.GetColumnSuggestions(header.SheetName, header.ColumnIndex, sampleRow)
                        : _workbookService.GetColumnSuggestions(header.ColumnIndex, sampleRow)
                    : [];

                var cellValue = rowValueOverrides[new CellKey(sheetName, sourceRow, header.ColumnIndex)];
                fieldDrafts.Add((header, sourceRow, sheetName, cellValue, dropdownOptions, dependency.ParentColumns, suggestionOptions));
            }

            var headerGroupIndices = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < fieldDrafts.Count; index++)
            {
                var (header, _, sheetName, _, _, _, _) = fieldDrafts[index];
                if (!header.Name.Contains('\n', StringComparison.Ordinal))
                {
                    continue;
                }

                var groupKey = BuildDropdownColorGroupKey(sheetName, header.Name);
                if (!headerGroupIndices.TryGetValue(groupKey, out var indices))
                {
                    indices = [];
                    headerGroupIndices[groupKey] = indices;
                }

                indices.Add(index);
            }

            var shownParentHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? lastSheetName = null;
            for (var fieldIndex = 0; fieldIndex < fieldDrafts.Count; fieldIndex++)
            {
                var (header, sourceRow, sheetName, cellValue, dropdownOptions, parentDropdownColumns, suggestionOptions) = fieldDrafts[fieldIndex];
                var rowBg = "Transparent";
                var rowBorder = "Transparent";
                var headerFg = "#111827";
                var inputBg = "#FFFFFF";
                var inputFg = "#111827";
                var groupBorderHex = "Transparent";
                var parentTitleBg = "Transparent";
                var parentTitleFg = "#111827";
                var colorGroupKey = BuildDropdownColorGroupKey(sheetName, header.Name);
                var hasDropdownColor = dropdownColorGroups.TryGetValue(colorGroupKey, out var paletteSlot);
                var isGroupedHeader = header.Name.Contains('\n', StringComparison.Ordinal);
                if (!isGroupedHeader && hasDropdownColor)
                {
                    var p = DropdownRowPalette[paletteSlot % DropdownRowPalette.Length];
                    rowBg = p.Bg;
                    rowBorder = p.Border;
                    headerFg = p.HeaderFg;
                    inputBg = p.InputBg;
                    inputFg = p.InputFg;
                }

                string headerDisplayName;
                string parentHeaderName = string.Empty;
                var showParentHeader = false;
                var isFirstInHeaderGroup = false;
                var isLastInHeaderGroup = false;
                if (isGroupedHeader)
                {
                    parentHeaderName = ExtractParentHeaderKey(header.Name);
                    headerDisplayName = ExtractChildHeaderName(header.Name);
                    if (string.IsNullOrWhiteSpace(headerDisplayName))
                    {
                        headerDisplayName = header.Name;
                    }

                    showParentHeader = shownParentHeaders.Add(colorGroupKey);
                    if (headerGroupIndices.TryGetValue(colorGroupKey, out var indices) && indices.Count > 0)
                    {
                        isFirstInHeaderGroup = indices[0] == fieldIndex;
                        isLastInHeaderGroup = indices[^1] == fieldIndex;
                    }

                    groupBorderHex = hasDropdownColor
                        ? DropdownRowPalette[paletteSlot % DropdownRowPalette.Length].Border
                        : "#D1D5DB";
                    if (showParentHeader && hasDropdownColor)
                    {
                        var palette = DropdownRowPalette[paletteSlot % DropdownRowPalette.Length];
                        parentTitleBg = palette.Bg;
                        parentTitleFg = palette.HeaderFg;
                    }
                }
                else
                {
                    headerDisplayName = header.Name;
                }

                var showSheetSeparator = IsMultiSheetMode
                    && !string.IsNullOrEmpty(sheetName)
                    && !string.Equals(lastSheetName, sheetName, StringComparison.Ordinal);
                if (!string.IsNullOrEmpty(sheetName))
                {
                    lastSheetName = sheetName;
                }

                var field = new EditableField
                {
                    HeaderName = header.Name,
                    HeaderDisplayName = headerDisplayName,
                    ParentHeaderName = parentHeaderName,
                    ShowParentHeader = showParentHeader,
                    IsGroupedUnderParentHeader = isGroupedHeader,
                    IsFirstInHeaderGroup = isFirstInHeaderGroup,
                    IsLastInHeaderGroup = isLastInHeaderGroup,
                    GroupBorderHex = groupBorderHex,
                    ParentTitleBackgroundHex = parentTitleBg,
                    ParentTitleForegroundHex = parentTitleFg,
                    ColumnIndex = header.ColumnIndex,
                    SheetName = sheetName,
                    ShowSheetSeparator = showSheetSeparator,
                    SheetSeparatorTitle = sheetName,
                    SourceRowIndex = sourceRow,
                    ParentDropdownColumns = parentDropdownColumns,
                    SuggestionOptions = suggestionOptions,
                    Value = cellValue,
                    RowHighlightBackgroundHex = rowBg,
                    RowHighlightBorderHex = rowBorder,
                    RowHeaderForegroundHex = headerFg,
                    RowInputBackgroundHex = inputBg,
                    RowInputForegroundHex = inputFg
                };
                field.ConfigureDropdownRole(parentDropdownColumns.Count > 0, dropdownOptions);
                field.SetInitialDropdownOptions(dropdownOptions);
                var dropdownSearchIndex = dropdownOptions.Count > DropdownLimits.SearchableThreshold
                    ? _workbookService.GetOrCreateDropdownSearchIndex(sheetName, sourceRow, header.ColumnIndex, rowValueOverrides)
                    : null;
                field.InitializeDropdownPresentation(dropdownSearchIndex);
                field.ValueChanged += EditableFieldOnValueChanged;
                field.DropdownSelectionCommitted += EditableFieldOnDropdownSelectionCommitted;
                EditableFields.Add(field);
            }
        }
        finally
        {
            _suppressAutosave = false;
        }

        OnPropertyChanged(nameof(CurrentRecordEntryCaption));
    }

    private void EditableFieldOnValueChanged(object? sender, EventArgs e)
    {
        if (_suppressAutosave || SelectedRecord is null || sender is not EditableField field)
        {
            return;
        }

        var pendingRow = IsMultiSheetMode ? field.SourceRowIndex : SelectedRecord.RowIndex;
        var pendingSheet = IsMultiSheetMode ? field.SheetName : (_workbookService.ActiveSheetName ?? string.Empty);

        if (_pendingRowIndex == -1)
        {
            _pendingRowIndex = pendingRow;
        }
        else if (_pendingRowIndex != pendingRow && !IsMultiSheetMode)
        {
            FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
            _pendingRowIndex = pendingRow;
        }

        var cellKey = new CellKey(pendingSheet, pendingRow, field.ColumnIndex);
        if (field.HasNonEmptyDropdownValue && field.IsDropdownValueInvalid)
        {
            _pendingCellValues.Remove(cellKey);
            StatusMessage = $"{field.HeaderDisplayName}: {field.DropdownValidationMessage}";
            return;
        }

        _pendingCellValues[cellKey] = field.Value;
        StatusMessage = IsAutoSaveEnabled
            ? $"Dòng {_pendingRowIndex}: thay đổi sẽ tự lưu file khi chuyển sang tên khác hoặc khi đóng ứng dụng."
            : $"Dòng {_pendingRowIndex} có thay đổi chưa lưu.";

    }

    private void EditableFieldOnDropdownSelectionCommitted(object? sender, EventArgs e)
    {
        if (_suppressAutosave || SelectedRecord is null || sender is not EditableField field)
        {
            return;
        }

        RefreshDependentDropdownFields(field);
    }

    private void RefreshDependentDropdownFields(EditableField changedField)
    {
        if (SelectedRecord is null)
        {
            return;
        }

        var hasDependents = EditableFields.Any(f =>
            f.IsDependentDropdown
            && f.ParentDropdownColumns.Contains(changedField.ColumnIndex)
            && IsSameEditableRow(f, changedField));
        if (!hasDependents)
        {
            return;
        }

        var overrides = BuildCurrentCellValueOverrides();
        EditableField? autoOpenField = null;
        _suppressAutosave = true;
        try
        {
            foreach (var dependentField in EditableFields
                         .Where(f => f.IsDependentDropdown
                             && f.ParentDropdownColumns.Contains(changedField.ColumnIndex)
                             && IsSameEditableRow(f, changedField))
                         .OrderBy(f => f.ColumnIndex))
            {
                var options = NormalizeDropdownOptions(_workbookService.GetDropdownOptions(
                    dependentField.SheetName,
                    dependentField.SourceRowIndex,
                    dependentField.ColumnIndex,
                    overrides));
                dependentField.ConfigureDropdownRole(isDependentColumn: true, options);
                var searchIndex = options.Count > DropdownLimits.SearchableThreshold
                    ? _workbookService.GetOrCreateDropdownSearchIndex(
                        dependentField.SheetName,
                        dependentField.SourceRowIndex,
                        dependentField.ColumnIndex,
                        overrides)
                    : null;
                dependentField.ReplaceDropdownOptions(options, searchIndex);

                if (autoOpenField is null && HasSelectableDropdownOptions(options))
                {
                    autoOpenField = dependentField;
                }

                if (!dependentField.IsDropdownValueAllowed(dependentField.Value))
                {
                    _suppressAutosave = true;
                    try
                    {
                        dependentField.Value = string.Empty;
                    }
                    finally
                    {
                        _suppressAutosave = false;
                    }

                    var cellKey = new CellKey(
                        dependentField.SheetName,
                        dependentField.SourceRowIndex,
                        dependentField.ColumnIndex);
                    _pendingCellValues[cellKey] = string.Empty;
                }
            }
        }
        finally
        {
            _suppressAutosave = false;
        }

        autoOpenField?.RequestOpenDropdown();
    }

    private static bool HasSelectableDropdownOptions(IReadOnlyList<string> options) =>
        options.Any(static value => !string.IsNullOrWhiteSpace(value));

    private Dictionary<CellKey, string> BuildCurrentCellValueOverrides()
    {
        var overrides = new Dictionary<CellKey, string>();
        foreach (var field in EditableFields)
        {
            var sheetName = IsMultiSheetMode
                ? field.SheetName
                : _workbookService.ActiveSheetName ?? string.Empty;
            overrides[new CellKey(sheetName, field.SourceRowIndex, field.ColumnIndex)] = field.Value;
        }

        return overrides;
    }

    private static bool IsSameEditableRow(EditableField left, EditableField right) =>
        left.SourceRowIndex == right.SourceRowIndex
        && string.Equals(left.SheetName, right.SheetName, StringComparison.Ordinal);

    private static IReadOnlyList<string> NormalizeDropdownOptions(IReadOnlyList<string> dropdownOptions)
    {
        if (dropdownOptions.Count == 0)
        {
            return dropdownOptions;
        }

        var normalized = dropdownOptions
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        normalized.Add(string.Empty);
        return normalized;
    }

    private void FlushPendingChanges(bool saveToDisk)
    {
        try
        {
            FlushPendingChangesCore(saveToDisk);
        }
        catch (Exception ex)
        {
            // Tránh crash khi autosave / ghi ô nếu có lỗi không lường trước.
            var root = GetInnermostException(ex);
            var detail = $"{root.GetType().Name}: {root.Message}";
            WarnSaveFailed(
                "Không hoàn tất lưu hoặc ghi dữ liệu. Nếu file đang mở ở ứng dụng khác, hãy đóng rồi thử lại.\n\n"
                + $"Chi tiết: {detail}");
            StatusMessage = detail;
        }
    }

    private static Exception GetInnermostException(Exception ex)
    {
        while (ex.InnerException is not null)
        {
            ex = ex.InnerException;
        }

        return ex;
    }

    private void FlushPendingChangesCore(bool saveToDisk)
    {
        if (TryGetInvalidDropdownField(out var invalidField))
        {
            var message = $"{invalidField.HeaderDisplayName}: {invalidField.DropdownValidationMessage}";
            StatusMessage = message;
            WarnSaveFailed(message + "\n\nHãy chọn giá trị hợp lệ trong dropdown trước khi lưu.");
            return;
        }

        if (_pendingRowIndex == -1 || _pendingCellValues.Count == 0)
        {
            if (saveToDisk && _workbookService.HasUnsavedChanges())
            {
                if (!_workbookService.TrySave(out var err))
                {
                    WarnSaveFailed(err);
                    StatusMessage = err ?? "Lưu thất bại.";
                }
                else
                {
                    StatusMessage = $"Đã lưu file lúc {DateTime.Now:HH:mm:ss}.";
                }
            }
            return;
        }

        foreach (var cell in _pendingCellValues)
        {
            var oldValue = string.IsNullOrEmpty(cell.Key.SheetName)
                ? _workbookService.ReadCellText(cell.Key.RowIndex, cell.Key.ColumnIndex)
                : _workbookService.ReadCellText(cell.Key.SheetName, cell.Key.RowIndex, cell.Key.ColumnIndex);
            var normalizedNewValue = cell.Value;
            if (string.IsNullOrEmpty(cell.Key.SheetName))
            {
                _workbookService.UpdateCellValue(cell.Key.RowIndex, cell.Key.ColumnIndex, normalizedNewValue);
                _workbookService.UpdateSuggestionCacheAfterCellEdit(cell.Key.RowIndex, cell.Key.ColumnIndex, SampleRowThreshold, oldValue, normalizedNewValue);
            }
            else
            {
                _workbookService.UpdateCellValue(cell.Key.SheetName, cell.Key.RowIndex, cell.Key.ColumnIndex, normalizedNewValue);
                var sampleRow = GetSampleRowForSheet(cell.Key.SheetName);
                _workbookService.UpdateSuggestionCacheAfterCellEdit(cell.Key.SheetName, cell.Key.RowIndex, cell.Key.ColumnIndex, sampleRow, oldValue, normalizedNewValue);
            }

            RefreshSuggestionFieldsForColumn(cell.Key.SheetName, cell.Key.ColumnIndex);
        }

        if (saveToDisk)
        {
            if (!_workbookService.TrySave(out var err))
            {
                WarnSaveFailed(err);
                StatusMessage = err ?? "Lưu thất bại.";
            }
            else
            {
                StatusMessage = $"Đã lưu dòng {_pendingRowIndex} lúc {DateTime.Now:HH:mm:ss}.";
            }
        }
        else
        {
            _workbookService.CommitPendingEditsToWorksheetOnly();
            StatusMessage = $"Dòng {_pendingRowIndex} có thay đổi chưa ghi file.";
        }
        _pendingCellValues.Clear();
        _pendingRowIndex = -1;
    }

    private bool TryGetInvalidDropdownField(out EditableField invalidField)
    {
        foreach (var field in EditableFields)
        {
            if (field.HasNonEmptyDropdownValue && field.IsDropdownValueInvalid)
            {
                invalidField = field;
                return true;
            }
        }

        invalidField = null!;
        return false;
    }

    private static void WarnSaveFailed(string? message)
    {
        MessageBox.Show(
            message ?? "Không thể ghi file Excel.",
            "Không lưu được",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private void EditRecord(RecordItem? record)
    {
        record ??= SelectedRecord;
        if (record is null || !_isWorkbookLoaded || GetActiveNameColumnIndex() <= 0)
        {
            return;
        }

        if (IsSameRecord(SelectedRecord, record))
        {
            FlushPendingChanges(saveToDisk: false);
        }

        var nameColumnIndex = GetActiveNameColumnIndex();
        string current;
        if (IsMultiSheetMode && _multiSheetSession?.NameSheetName is { } nameSheetName)
        {
            current = _workbookService.ReadCellText(nameSheetName, record.RowIndex, nameColumnIndex);
        }
        else
        {
            current = _workbookService.ReadCellText(record.RowIndex, nameColumnIndex);
        }

        var input = ShowInputDialog("Tên mới (cột tên):", current, "Sửa tên dòng");
        if (string.IsNullOrEmpty(input))
        {
            StatusMessage = "Đã hủy sửa tên.";
            return;
        }

        if (string.Equals(input, current, StringComparison.Ordinal))
        {
            StatusMessage = "Không đổi (tên giữ nguyên).";
            return;
        }

        if (IsMultiSheetMode && _multiSheetSession?.NameSheetName is { } sheetName)
        {
            var sampleRow = GetSampleRowForSheet(sheetName);
            _workbookService.UpdateCellValue(sheetName, record.RowIndex, nameColumnIndex, input);
            _workbookService.UpdateSuggestionCacheAfterCellEdit(sheetName, record.RowIndex, nameColumnIndex, sampleRow, current, input);
            RefreshSuggestionFieldsForColumn(sheetName, nameColumnIndex);
        }
        else
        {
            _workbookService.UpdateCellValue(record.RowIndex, nameColumnIndex, input);
            _workbookService.UpdateSuggestionCacheAfterCellEdit(record.RowIndex, nameColumnIndex, SampleRowThreshold, current, input);
            RefreshSuggestionFieldsForColumn(nameColumnIndex);
        }

        ReloadRecordItemsFromWorkbook();
        SelectedRecord = IsMultiSheetMode
            ? Records.FirstOrDefault(r => r.LogicalIndex == record.LogicalIndex)
            : Records.FirstOrDefault(r => r.RowIndex == record.RowIndex);

        if (IsAutoSaveEnabled)
        {
            if (!_workbookService.TrySave(out var err))
            {
                WarnSaveFailed(err);
                StatusMessage = err ?? "Đã đổi tên nhưng chưa ghi được file.";
                return;
            }
        }
        else
        {
            _workbookService.CommitPendingEditsToWorksheetOnly();
        }

        StatusMessage = $"Đã đổi tên dòng thành '{input}'.";
    }

    private void DeleteRecord(RecordItem? record)
    {
        record ??= SelectedRecord;
        if (record is null || !_isWorkbookLoaded || GetActiveNameColumnIndex() <= 0)
        {
            return;
        }

        var confirm = MessageBox.Show(
            $"Xóa dòng \"{record.KeyDisplay}\" (Excel hàng {record.RowIndex})?\nKhông hoàn tác.",
            "Xác nhận xóa",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning,
            MessageBoxResult.Cancel);
        if (confirm != MessageBoxResult.OK)
        {
            return;
        }

        var listBefore = Records.ToList();
        var indexBefore = IsMultiSheetMode
            ? listBefore.FindIndex(r => r.LogicalIndex == record.LogicalIndex)
            : listBefore.FindIndex(r => r.RowIndex == record.RowIndex);
        if (indexBefore < 0)
        {
            return;
        }

        if (IsSameRecord(SelectedRecord, record))
        {
            FlushPendingChanges(saveToDisk: false);
        }

        try
        {
            if (IsMultiSheetMode && _multiSheetSession is not null)
            {
                foreach (var sheetCfg in _multiSheetSession.Sheets)
                {
                    var rowToDelete = sheetCfg.FirstDataRow + record.LogicalIndex;
                    _workbookService.DeleteRowOnSheet(sheetCfg.SheetName, rowToDelete);
                }
            }
            else
            {
                _workbookService.DeleteRow(record.RowIndex);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không xóa được dòng: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        ReloadRecordItemsFromWorkbook();

        if (Records.Count == 0)
        {
            SelectedRecord = null;
        }
        else
        {
            var pickIndex = Math.Min(indexBefore, Records.Count - 1);
            SelectedRecord = Records[pickIndex];
        }

        if (IsAutoSaveEnabled)
        {
            if (!_workbookService.TrySave(out var err))
            {
                WarnSaveFailed(err);
                StatusMessage = err ?? "Đã xóa dòng nhưng chưa ghi được file.";
                return;
            }
        }
        else
        {
            _workbookService.CommitPendingEditsToWorksheetOnly();
        }

        StatusMessage = $"Đã xóa dòng (Excel hàng {record.RowIndex}).";
    }

    private bool FilterRecords(object obj)
    {
        if (obj is not RecordItem record)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        return VietnameseTextHelper.ContainsNormalized(record.KeyDisplay, SearchText)
            || VietnameseTextHelper.ContainsNormalized(record.TooltipPreview, SearchText);
    }

    private void RememberHiddenHeaders()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        var hiddenHeaders = HeaderColumns.Where(h => !h.IsVisible).Select(h => h.Name).ToList();
        _profileStore.SaveHiddenHeaders(_profileKey, hiddenHeaders);
        StatusMessage = "Đã tự ghi nhớ danh sách mục nhập liệu ẩn.";
    }

    private void SaveNow()
    {
        FlushPendingChanges(saveToDisk: true);
    }

    private void SaveAs()
    {
        if (!_isWorkbookLoaded)
        {
            return;
        }

        FlushPendingChanges(saveToDisk: false);
        var dialog = new SaveFileDialog
        {
            Title = "Save As",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = Path.GetFileName(_workbookService.WorkbookPath)
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        if (!_workbookService.TrySaveAs(dialog.FileName, out var err))
        {
            WarnSaveFailed(err);
            StatusMessage = err ?? "Không lưu được file mới.";
            return;
        }

        FilePath = dialog.FileName;
        StatusMessage = $"Đã lưu file mới: {Path.GetFileName(dialog.FileName)}";
    }

    public string? ExportHssk(
        string templatePath,
        string outputFolder,
        IReadOnlyCollection<int> selectedRowIndexes,
        string fileNote,
        bool exportToSingleFile,
        string singleFileName,
        int genderColumnIndex,
        IReadOnlyCollection<int> femaleOnlyColumnIndexes,
        Action<int, int>? progressCallback = null)
    {
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel.";
        }

        if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
        {
            return "File mẫu Word không hợp lệ.";
        }

        if (selectedRowIndexes.Count == 0)
        {
            return "Chưa chọn dòng để xuất.";
        }

        var destination = string.IsNullOrWhiteSpace(outputFolder)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "HSSK_Export")
            : outputFolder.Trim();

        try
        {
            Directory.CreateDirectory(destination);
            FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
            var sanitizedNote = SanitizeFileName(fileNote);

            var columns = HeaderColumns.Select(h => h.ColumnIndex).Distinct().ToList();
            if (NameColumnIndex > 0 && !columns.Contains(NameColumnIndex))
            {
                columns.Add(NameColumnIndex);
            }

            var sampleRowValues = SampleRowThreshold > 0
                ? _workbookService.ReadRowValues(SampleRowThreshold, columns)
                : new Dictionary<int, string>();

            var femaleOnlyColumns = femaleOnlyColumnIndexes.Count > 0
                ? femaleOnlyColumnIndexes.ToHashSet()
                : [];

            var selectedRows = selectedRowIndexes.Distinct().ToList();
            var resolvedRows = new List<IReadOnlyDictionary<int, string>>();
            var resolvedNamedPlaceholders = new List<IReadOnlyDictionary<string, string>>();
            var exportCount = 0;
            foreach (var rowIndex in selectedRows)
            {
                var rowValues = _workbookService.ReadRowValues(rowIndex, columns);
                var genderText = genderColumnIndex > 0
                    ? rowValues.TryGetValue(genderColumnIndex, out var genderValue) ? genderValue : string.Empty
                    : string.Empty;
                var useFemaleOnlyColumns = HsskExportGenderHelper.ShouldUseFemaleOnlyColumns(genderText);

                foreach (var columnIndex in columns)
                {
                    if (!rowValues.TryGetValue(columnIndex, out var currentValue) || !string.IsNullOrWhiteSpace(currentValue))
                    {
                        continue;
                    }

                    if (!useFemaleOnlyColumns && femaleOnlyColumns.Contains(columnIndex))
                    {
                        continue;
                    }

                    if (sampleRowValues.TryGetValue(columnIndex, out var sampleValue) && !string.IsNullOrWhiteSpace(sampleValue))
                    {
                        rowValues[columnIndex] = sampleValue;
                    }
                }

                resolvedRows.Add(rowValues);
                var phanLoai = ResolveMaxLoaiLabelForRow(rowIndex);
                var namedPlaceholders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["phanloai"] = phanLoai
                };
                resolvedNamedPlaceholders.Add(namedPlaceholders);
                var personName = rowValues.TryGetValue(NameColumnIndex, out var rawName) ? rawName : string.Empty;
                var cleanName = SanitizeFileName(personName);
                if (string.IsNullOrWhiteSpace(cleanName))
                {
                    cleanName = $"Row_{rowIndex}";
                }

                if (exportToSingleFile)
                {
                    continue;
                }

                var fileBaseName = string.IsNullOrWhiteSpace(sanitizedNote)
                    ? cleanName
                    : $"{cleanName}_{sanitizedNote}";
                var outputPath = BuildUniqueOutputPath(destination, fileBaseName);
                WordTemplateExportService.ExportFromTemplate(templatePath, outputPath, rowValues, namedPlaceholders);
                exportCount++;
                progressCallback?.Invoke(exportCount, selectedRows.Count);
            }

            if (exportToSingleFile)
            {
                var combinedBase = SanitizeFileName(singleFileName);
                if (string.IsNullOrWhiteSpace(combinedBase))
                {
                    combinedBase = "HSSK_TongHop";
                }

                if (!string.IsNullOrWhiteSpace(sanitizedNote))
                {
                    combinedBase = $"{combinedBase}_{sanitizedNote}";
                }

                var outputPath = BuildUniqueOutputPath(destination, combinedBase);
                WordTemplateExportService.ExportManyRecordsToSingleFile(
                    templatePath,
                    outputPath,
                    resolvedRows,
                    resolvedNamedPlaceholders,
                    progressCallback);
                StatusMessage = $"Đã xuất {resolvedRows.Count} hồ sơ vào 1 file: {outputPath}";
            }
            else
            {
                StatusMessage = $"Đã xuất {exportCount} file HSSK vào: {destination}";
            }

            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể xuất HSSK: {root.Message}";
        }
    }

    public ImportFileExportPreview? TryGetImportFileExportPreview()
    {
        if (!_isWorkbookLoaded || HeaderColumns.Count == 0)
        {
            return null;
        }

        if (!TryBuildImportFileExportPlans(out var plans, out var preview))
        {
            return null;
        }

        return preview;
    }

    public string? ExportImportFile(string outputPath)
    {
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel.";
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return "Chưa chọn file xuất.";
        }

        if (HeaderColumns.Count == 0)
        {
            return "Chưa có tiêu đề cột.";
        }

        if (TryGetInvalidDropdownField(out var invalidField))
        {
            return $"{invalidField.HeaderDisplayName}: {invalidField.DropdownValidationMessage}";
        }

        try
        {
            FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
            if (!TryBuildImportFileExportPlans(out var plans, out var preview))
            {
                return "Không xác định được phạm vi dòng tên để xuất.";
            }

            var result = ImportFileExportService.Export(_workbookService.WorkbookPath, outputPath, plans);
            StatusMessage =
                $"Đã xuất file import: {Path.GetFileName(result.OutputPath)} "
                + $"(bổ sung {result.FilledCellCount} ô, {result.ProcessedRowCount} dòng, "
                + $"đến hàng tên {preview.LastNameRow}).";
            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể xuất file import: {root.Message}";
        }
    }

    private (int GenderColumnIndex, string? GenderSheetName, int GenderSheetFirstDataRow) ResolveImportFileGenderContext(
        string planSheetName,
        int planFirstDataRow)
    {
        var genderColumn = GetImportFileExportProfile().GenderColumnIndex;
        if (genderColumn <= 0)
        {
            return (0, null, 0);
        }

        var genderHeader = HeaderColumns.FirstOrDefault(h => h.ColumnIndex == genderColumn);
        if (IsMultiSheetMode && _multiSheetSession?.NameSheet is { } nameSheet)
        {
            var genderSheetName = genderHeader?.SheetName ?? nameSheet.SheetName;
            var genderSheetCfg = _multiSheetSession.Sheets.FirstOrDefault(s =>
                string.Equals(s.SheetName, genderSheetName, StringComparison.Ordinal));
            var genderFirstDataRow = genderSheetCfg?.FirstDataRow ?? nameSheet.FirstDataRow;
            return (genderColumn, genderSheetName, genderFirstDataRow);
        }

        var genderSheet = genderHeader?.SheetName ?? planSheetName;
        return (genderColumn, genderSheet, planFirstDataRow);
    }

    private bool TryBuildImportFileExportPlans(
        out List<ImportFileSheetPlan> plans,
        out ImportFileExportPreview preview)
    {
        plans = [];
        preview = new ImportFileExportPreview();
        var importProfile = GetImportFileExportProfile();

        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            var nameSheet = _multiSheetSession.NameSheet;
            if (nameSheet is null)
            {
                return false;
            }

            var nameColumnIndex = GetActiveNameColumnIndex();
            var firstDataRow = nameSheet.FirstDataRow;
            var lastNameRow = _workbookService.FindLastNameRow(
                nameSheet.SheetName,
                nameColumnIndex,
                firstDataRow);
            if (lastNameRow < firstDataRow)
            {
                return false;
            }

            var rowSpan = lastNameRow - firstDataRow;
            foreach (var sheetCfg in _multiSheetSession.Sheets)
            {
                var columns = HeaderColumns
                    .Where(h => string.Equals(h.SheetName, sheetCfg.SheetName, StringComparison.Ordinal))
                    .Select(h => h.ColumnIndex)
                    .Distinct()
                    .ToList();
                if (string.Equals(sheetCfg.SheetName, nameSheet.SheetName, StringComparison.Ordinal)
                    && nameColumnIndex > 0
                    && !columns.Contains(nameColumnIndex))
                {
                    columns.Add(nameColumnIndex);
                }

                if (columns.Count == 0)
                {
                    continue;
                }

                var sampleRow = GetSampleRowForSheet(sheetCfg.SheetName);
                var genderContext = ResolveImportFileGenderContext(sheetCfg.SheetName, sheetCfg.FirstDataRow);
                plans.Add(new ImportFileSheetPlan
                {
                    SheetName = sheetCfg.SheetName,
                    FirstDataRow = sheetCfg.FirstDataRow,
                    LastDataRow = sheetCfg.FirstDataRow + rowSpan,
                    SampleRow = sampleRow,
                    ColumnIndexes = columns,
                    SkipSampleColumnIndexes = importProfile.SkipSampleColumnIndexes,
                    GenderColumnIndex = genderContext.GenderColumnIndex,
                    GenderSheetName = genderContext.GenderSheetName,
                    GenderSheetFirstDataRow = genderContext.GenderSheetFirstDataRow
                });
            }

            if (plans.Count == 0)
            {
                return false;
            }

            preview = new ImportFileExportPreview
            {
                IsMultiSheetMode = true,
                NameSheetName = nameSheet.SheetName,
                FirstDataRow = firstDataRow,
                LastNameRow = lastNameRow,
                SampleRow = GetSampleRowForSheet(nameSheet.SheetName),
                SheetCount = plans.Count
            };
            return true;
        }

        var singleFirstDataRow = HeaderRowNumber + 1;
        var singleNameColumn = NameColumnIndex;
        var singleLastNameRow = _workbookService.FindLastNameRow(
            _workbookService.ActiveSheetName ?? SelectedSheet,
            singleNameColumn,
            singleFirstDataRow);
        if (singleLastNameRow < singleFirstDataRow)
        {
            return false;
        }

        var singleColumns = HeaderColumns.Select(h => h.ColumnIndex).Distinct().ToList();
        if (singleNameColumn > 0 && !singleColumns.Contains(singleNameColumn))
        {
            singleColumns.Add(singleNameColumn);
        }

        var singleSampleRow = SampleRowThreshold > 0 ? SampleRowThreshold : singleFirstDataRow;
        var singleSheetName = _workbookService.ActiveSheetName ?? SelectedSheet;
        var singleGenderContext = ResolveImportFileGenderContext(singleSheetName, singleFirstDataRow);
        plans.Add(new ImportFileSheetPlan
        {
            SheetName = singleSheetName,
            FirstDataRow = singleFirstDataRow,
            LastDataRow = singleLastNameRow,
            SampleRow = singleSampleRow,
            ColumnIndexes = singleColumns,
            SkipSampleColumnIndexes = importProfile.SkipSampleColumnIndexes,
            GenderColumnIndex = singleGenderContext.GenderColumnIndex,
            GenderSheetName = singleGenderContext.GenderSheetName,
            GenderSheetFirstDataRow = singleGenderContext.GenderSheetFirstDataRow
        });

        preview = new ImportFileExportPreview
        {
            IsMultiSheetMode = false,
            NameSheetName = _workbookService.ActiveSheetName ?? SelectedSheet,
            FirstDataRow = singleFirstDataRow,
            LastNameRow = singleLastNameRow,
            SampleRow = singleSampleRow,
            SheetCount = 1
        };
        return true;
    }

    private string ResolveMaxLoaiLabelForRow(int rowIndex)
    {
        var texts = _workbookService.ReadAllCellTextsInRow(rowIndex);
        var maxLoai = 0;
        foreach (var cellText in texts)
        {
            if (string.IsNullOrWhiteSpace(cellText))
            {
                continue;
            }

            foreach (Match match in LoaiRegex.Matches(cellText))
            {
                if (!int.TryParse(match.Groups[1].Value, out var value))
                {
                    continue;
                }

                if (value > maxLoai)
                {
                    maxLoai = value;
                }
            }
        }

        return maxLoai > 0 ? $"Loại {maxLoai}" : "Loại 1";
    }

    public string? ExportHeaderPlaceholdersToExcel(string outputPath)
    {
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel dữ liệu.";
        }

        if (HeaderColumns.Count == 0)
        {
            return "Chưa có tiêu đề để xuất.";
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return "Đường dẫn file xuất không hợp lệ.";
        }

        try
        {
            var fileInfo = new FileInfo(outputPath);
            if (fileInfo.Directory is not null)
            {
                Directory.CreateDirectory(fileInfo.Directory.FullName);
            }

            ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Mapping");

            sheet.Cells[1, 1].Value = "Placeholder";
            sheet.Cells[1, 2].Value = "Số cột";
            sheet.Cells[1, 3].Value = "Tiêu đề";

            var row = 2;
            foreach (var header in HeaderColumns.OrderBy(h => h.ColumnIndex))
            {
                sheet.Cells[row, 1].Value = $"${header.ColumnIndex}";
                sheet.Cells[row, 2].Value = header.ColumnIndex;
                sheet.Cells[row, 3].Value = header.Name;
                row++;
            }

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            package.SaveAs(fileInfo);
            StatusMessage = $"Đã xuất mapping tiêu đề ra file: {outputPath}";
            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể xuất mapping tiêu đề: {root.Message}";
        }
    }

    private void AddRecord()
    {
        if (!_isWorkbookLoaded || HeaderColumns.Count == 0 || GetActiveNameColumnIndex() <= 0)
        {
            return;
        }

        var nameColumnIndex = GetActiveNameColumnIndex();
        var nameValue = ShowInputDialog("Nhập tên cho dòng mới:", string.Empty, "Thêm dòng mới");
        if (string.IsNullOrWhiteSpace(nameValue))
        {
            StatusMessage = "Đã hủy thêm dòng mới.";
            return;
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        int newRowIndex;
        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            newRowIndex = -1;
            foreach (var sheetCfg in _multiSheetSession.Sheets)
            {
                var sheetHeaders = HeaderColumns
                    .Where(h => string.Equals(h.SheetName, sheetCfg.SheetName, StringComparison.Ordinal))
                    .Select(h => new HeaderCell(h.ColumnIndex, h.Name))
                    .ToList();
                var isNameSheet = string.Equals(sheetCfg.SheetName, _multiSheetSession.NameSheetName, StringComparison.Ordinal);
                var addedRow = _workbookService.AddRecordAtSheetLogicalEnd(
                    sheetCfg.SheetName,
                    sheetCfg.FirstDataRow,
                    isNameSheet ? nameColumnIndex : 0,
                    isNameSheet ? nameValue : string.Empty,
                    sheetHeaders);
                if (isNameSheet)
                {
                    newRowIndex = addedRow;
                }
            }
        }
        else
        {
            newRowIndex = _workbookService.AddRecordAtLogicalEnd(
                HeaderRowNumber,
                NameColumnIndex,
                nameValue,
                HeaderColumns.Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
        }

        if (IsAutoSaveEnabled)
        {
            if (!_workbookService.TrySave(out var err))
            {
                WarnSaveFailed(err);
                StatusMessage = err ?? "Đã thêm dòng nhưng chưa ghi được file.";
                LoadRecords();
                SelectedRecord = IsMultiSheetMode
                    ? Records.LastOrDefault()
                    : Records.FirstOrDefault(r => r.RowIndex == newRowIndex) ?? Records.LastOrDefault();
                return;
            }
        }
        else
        {
            _workbookService.CommitPendingEditsToWorksheetOnly();
        }

        LoadRecords();
        SelectedRecord = IsMultiSheetMode
            ? Records.LastOrDefault()
            : Records.FirstOrDefault(r => r.RowIndex == newRowIndex) ?? Records.LastOrDefault();
        StatusMessage = IsAutoSaveEnabled
            ? $"Đã thêm dòng mới '{nameValue}' tại dòng {newRowIndex}."
            : $"Đã thêm dòng mới '{nameValue}' tại dòng {newRowIndex} (chưa ghi file).";
    }

    public IReadOnlyList<string> GetDistinctSortColumnValues(HeaderDefinition sortHeader)
    {
        return _workbookService.GetDistinctColumnValues(
            ResolveSingleSheetSortStartRow(),
            sortHeader.ColumnIndex,
            ResolveSingleSheetSortExcludedRow());
    }

    public string? SortRecordsByColumn(
        HeaderDefinition sortHeader,
        ColumnSortMode mode,
        IReadOnlyList<string>? customOrder = null)
    {
        if (!_isWorkbookLoaded || NameColumnIndex <= 0)
        {
            return "Chưa mở file Excel hoặc chưa cấu hình cột tên.";
        }

        if (HeaderColumns.Count == 0)
        {
            return "Chưa có tiêu đề cột để sắp xếp.";
        }

        if (mode == ColumnSortMode.Custom && (customOrder is null || customOrder.Count == 0))
        {
            return "Chưa có thứ tự tùy chỉnh để sắp xếp.";
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);

        try
        {
            _workbookService.SortRowsByColumn(
                ResolveSingleSheetSortStartRow(),
                sortHeader.ColumnIndex,
                mode,
                customOrder,
                ResolveSingleSheetSortExcludedRow());
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể sắp xếp dòng: {root.Message}";
        }

        var modeLabel = mode switch
        {
            ColumnSortMode.Descending => "Z → A",
            ColumnSortMode.Custom => "tùy chỉnh",
            _ => "A → Z"
        };

        if (IsAutoSaveEnabled)
        {
            if (!_workbookService.TrySave(out var err))
            {
                WarnSaveFailed(err);
                StatusMessage = err ?? "Đã sắp xếp nhưng chưa ghi được file.";
            }
            else
            {
                StatusMessage = $"Đã sắp xếp dòng theo cột '{sortHeader.Name}' ({modeLabel}).";
            }
        }
        else
        {
            _workbookService.CommitPendingEditsToWorksheetOnly();
            StatusMessage = $"Đã sắp xếp dòng theo cột '{sortHeader.Name}' ({modeLabel}, chưa ghi file).";
        }

        LoadRecords();
        SelectedRecord = Records.FirstOrDefault();
        return null;
    }

    public DataLinkProfile? TryLoadDataLinkProfile()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return null;
        }

        return _dataLinkProfileStore.TryLoad(_profileKey);
    }

    public void SaveDataLinkProfile(DataLinkProfile profile)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        _dataLinkProfileStore.Save(_profileKey, profile);
    }

    public IReadOnlyList<string> GetSavedMappingProfileNames()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return [];
        }

        return _dataLinkProfileStore.GetSavedProfileNames(_profileKey);
    }

    public DataLinkProfile? TryLoadNamedMappingProfile(string name)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return null;
        }

        return _dataLinkProfileStore.TryLoadNamed(_profileKey, name);
    }

    public void SaveNamedMappingProfile(string name, DataLinkProfile profile)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        _dataLinkProfileStore.SaveNamed(_profileKey, name, profile);
        _dataLinkProfileStore.SaveLastNamedProfileName(_profileKey, name);
    }

    public string? TryLoadLastNamedMappingProfileName()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return null;
        }

        return _dataLinkProfileStore.TryLoadLastNamedProfileName(_profileKey);
    }

    public void RememberLastNamedMappingProfile(string name)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        _dataLinkProfileStore.SaveLastNamedProfileName(_profileKey, name);
    }

    public void DeleteNamedMappingProfile(string name)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        _dataLinkProfileStore.DeleteNamed(_profileKey, name);
        if (string.Equals(TryLoadLastNamedMappingProfileName(), name, StringComparison.OrdinalIgnoreCase))
        {
            _dataLinkProfileStore.ClearLastNamedProfileName(_profileKey);
        }
    }

    public string? BackupNamedMappingProfiles(string outputPath)
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return "Chưa có ngữ cảnh profile để backup.";
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return "Đường dẫn file backup không hợp lệ.";
        }

        try
        {
            _dataLinkProfileStore.BackupNamedProfiles(_profileKey, outputPath);
            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể backup profile: {root.Message}";
        }
    }

    public string? RestoreNamedMappingProfiles(string inputPath, out IReadOnlyList<string> restoredNames)
    {
        restoredNames = [];
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return "Chưa có ngữ cảnh profile để restore.";
        }

        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            return "File backup profile không tồn tại.";
        }

        try
        {
            restoredNames = _dataLinkProfileStore.RestoreNamedProfiles(_profileKey, inputPath);
            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể restore profile: {root.Message}";
        }
    }

    public IReadOnlyList<string> GetDataFileSheetNames(string path)
    {
        ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
        using var package = new ExcelPackage(new FileInfo(path));
        return package.Workbook.Worksheets.Select(s => s.Name).ToList();
    }

    public IReadOnlyList<HeaderCell> LoadDataFileHeaders(string path, string sheetName, int headerRow)
    {
        using var reader = new ExcelLookupReader();
        reader.Open(path, sheetName);
        return reader.ReadHeaderRow(headerRow);
    }

    public IReadOnlyList<ImportListRowOption> LoadImportListRowOptions(ImportListProfile profile)
    {
        if (!_isWorkbookLoaded
            || string.IsNullOrWhiteSpace(profile.DataFilePath)
            || !File.Exists(profile.DataFilePath)
            || string.IsNullOrWhiteSpace(profile.DataSheetName)
            || profile.DataNameColumnIndex <= 0)
        {
            return [];
        }

        var dataFirstDataRow = Math.Max(profile.DataHeaderRow + 1, 1);

        using var reader = new ExcelLookupReader();
        reader.Open(profile.DataFilePath, profile.DataSheetName);
        var headers = LoadDataFileHeaders(profile.DataFilePath, profile.DataSheetName, profile.DataHeaderRow);
        var columnIndexes = headers.Select(h => h.ColumnIndex).Distinct().OrderBy(c => c).ToList();

        var options = new List<ImportListRowOption>();
        for (var row = dataFirstDataRow; row <= reader.GetEndRow(); row++)
        {
            var cells = new Dictionary<int, string>();
            var searchParts = new List<string>();
            foreach (var columnIndex in columnIndexes)
            {
                var cellText = reader.ReadCellText(row, columnIndex);
                cells[columnIndex] = cellText;
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    searchParts.Add(cellText);
                }
            }

            if (!cells.TryGetValue(profile.DataNameColumnIndex, out var name)
                || string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var option = new ImportListRowOption
            {
                DataRowIndex = row,
                KeyValue = name,
                NormalizedKey = VietnameseTextHelper.ToTonePlacementKey(name),
                CellValuesByColumn = cells,
                RowSearchText = string.Join(' ', searchParts),
                Status = ImportListMatchStatus.Ready
            };
            option.IsSelected = option.CanSelect;
            options.Add(option);
        }

        return options
            .OrderBy(o => o.KeyValue, StringComparer.OrdinalIgnoreCase)
            .ThenBy(o => o.DataRowIndex)
            .ToList();
    }

    public ImportListProfile? TryLoadImportListProfile() =>
        TryLoadDataLinkProfile() is { } profile ? ToImportListProfile(profile) : null;

    public void SaveImportListProfile(ImportListProfile profile) =>
        SaveDataLinkProfile(ToDataLinkProfileForStore(profile));

    public IReadOnlyList<string> GetSavedImportListProfileNames() => GetSavedMappingProfileNames();

    public ImportListProfile? TryLoadNamedImportListProfile(string name) =>
        TryLoadNamedMappingProfile(name) is { } profile ? ToImportListProfile(profile) : null;

    public void SaveNamedImportListProfile(string name, ImportListProfile profile) =>
        SaveNamedMappingProfile(name, ToDataLinkProfileForStore(profile));

    private static ImportListProfile ToImportListProfile(DataLinkProfile profile) => new()
    {
        DataFilePath = profile.DataFilePath,
        DataSheetName = profile.DataSheetName,
        DataHeaderRow = profile.DataHeaderRow,
        DataNameColumnIndex = profile.DataKeyColumnIndex,
        ColumnMappings = profile.ColumnMappings.Select(m => new ColumnMappingPair
        {
            TargetColumnIndex = m.TargetColumnIndex,
            TargetSheetName = m.TargetSheetName,
            SourceColumnIndex = m.SourceColumnIndex
        }).ToList()
    };

    private static DataLinkProfile ToDataLinkProfileForStore(ImportListProfile profile) => new()
    {
        DataFilePath = profile.DataFilePath,
        DataSheetName = profile.DataSheetName,
        DataHeaderRow = profile.DataHeaderRow,
        DataKeyColumnIndex = profile.DataNameColumnIndex,
        ColumnMappings = profile.ColumnMappings
    };

    public string? RunImportListLink(ImportListLinkRequest request, out DataLinkResult? result)
    {
        result = null;
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel dữ liệu.";
        }

        if (HeaderColumns.Count == 0)
        {
            return "Chưa có tiêu đề cột trên file gốc.";
        }

        var profile = request.Profile;
        if (string.IsNullOrWhiteSpace(profile.DataFilePath) || !File.Exists(profile.DataFilePath))
        {
            return "File danh sách không tồn tại.";
        }

        if (string.IsNullOrWhiteSpace(profile.DataSheetName))
        {
            return "Chưa chọn sheet file danh sách.";
        }

        if (profile.DataNameColumnIndex <= 0)
        {
            return "Chưa chọn cột tên trên file danh sách.";
        }

        if (profile.ColumnMappings.Count == 0)
        {
            return "Chưa có cặp cột ánh xạ nào.";
        }

        if (request.SelectedDataRowIndices.Count == 0)
        {
            return "Chưa chọn tên nào để import.";
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);

        if (!TryBuildImportMainContext(out var mainContext, out var contextError))
        {
            return contextError;
        }

        var formatReferenceRow = SampleRowThreshold > 0 ? SampleRowThreshold : Math.Max(1, HeaderRowNumber);
        var selectedRowSet = request.SelectedDataRowIndices.ToHashSet();

        try
        {
            using var reader = new ExcelLookupReader();
            reader.Open(profile.DataFilePath, profile.DataSheetName);

            var rowsCreated = 0;
            var cellsWritten = 0;
            var skippedEmpty = 0;

            foreach (var dataRowIndex in request.SelectedDataRowIndices.OrderBy(r => r))
            {
                if (!selectedRowSet.Remove(dataRowIndex))
                {
                    continue;
                }

                var name = reader.ReadCellText(dataRowIndex, profile.DataNameColumnIndex);
                if (string.IsNullOrWhiteSpace(name))
                {
                    skippedEmpty++;
                    continue;
                }

                var mainRow = AppendImportRowAtEnd(name.Trim());
                rowsCreated++;
                cellsWritten += _dataLinkService.WriteImportRowFromList(
                    _workbookService,
                    reader,
                    mainContext,
                    dataRowIndex,
                    mainRow,
                    profile.ColumnMappings,
                    formatReferenceRow);
            }

            if (rowsCreated == 0)
            {
                return "Không import được dòng nào (tên trống trong các dòng đã chọn).";
            }

            result = new DataLinkResult
            {
                RowsMatched = rowsCreated,
                RowsCreatedInMain = rowsCreated,
                CellsWritten = cellsWritten,
                SkippedSelectedNotInMain = skippedEmpty
            };

            SaveImportListProfile(profile);

            if (IsAutoSaveEnabled)
            {
                if (!_workbookService.TrySave(out var err))
                {
                    WarnSaveFailed(err);
                    StatusMessage = err ?? "Đã import nhưng chưa ghi được file.";
                }
                else
                {
                    StatusMessage = BuildImportListLinkStatusMessage(result);
                }
            }
            else
            {
                _workbookService.CommitPendingEditsToWorksheetOnly();
                StatusMessage = BuildImportListLinkStatusMessage(result) + " (chưa ghi file).";
            }

            var selectedLogical = SelectedRecord?.LogicalIndex;
            var selectedRow = SelectedRecord?.RowIndex;
            LoadRecords();
            if (IsMultiSheetMode && selectedLogical is int logicalIdx)
            {
                SelectedRecord = Records.FirstOrDefault(r => r.LogicalIndex == logicalIdx) ?? Records.FirstOrDefault();
            }
            else if (selectedRow is not null)
            {
                SelectedRecord = Records.FirstOrDefault(r => r.RowIndex == selectedRow) ?? Records.FirstOrDefault();
            }
            else
            {
                SelectedRecord = Records.FirstOrDefault();
            }

            if (SelectedRecord is not null)
            {
                LoadEditableFieldsForSelectedRecord();
            }

            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể import danh sách: {root.Message}";
        }
    }

    public string BuildImportListLinkSummary(DataLinkResult result)
    {
        var lines = new List<string>
        {
            $"Đã thêm {result.RowsCreatedInMain} dòng mới ở cuối file gốc.",
            $"Đã ghi {result.CellsWritten} ô theo ánh xạ cột."
        };

        if (result.SkippedSelectedNotInMain > 0)
        {
            lines.Add($"Bỏ qua {result.SkippedSelectedNotInMain} dòng đã chọn vì tên trống.");
        }

        return string.Join("\n", lines);
    }

    private static string BuildImportListLinkStatusMessage(DataLinkResult result) =>
        $"Đã import {result.RowsCreatedInMain} dòng, ghi {result.CellsWritten} ô.";

    private bool TryBuildImportMainContext(out DataLinkMainContext mainContext, out string? errorMessage)
    {
        mainContext = new DataLinkMainContext();
        errorMessage = null;

        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            var nameSheetName = _multiSheetSession.NameSheetName;
            var nameSheetCfg = _multiSheetSession.GetSheetConfig(nameSheetName);
            if (nameSheetCfg is null)
            {
                errorMessage = $"Không tìm thấy cấu hình sheet tên '{nameSheetName}'.";
                return false;
            }

            mainContext = new DataLinkMainContext
            {
                KeySheetName = nameSheetName,
                KeySheetFirstDataRow = nameSheetCfg.FirstDataRow,
                FirstDataRowBySheet = _multiSheetSession.Sheets.ToDictionary(
                    s => s.SheetName,
                    s => s.FirstDataRow,
                    StringComparer.Ordinal),
                FormatReferenceRowBySheet = _multiSheetSession.Sheets.ToDictionary(
                    s => s.SheetName,
                    s => GetSampleRowForSheet(s.SheetName),
                    StringComparer.Ordinal)
            };
            return true;
        }

        mainContext = new DataLinkMainContext
        {
            KeySheetName = string.Empty,
            KeySheetFirstDataRow = ResolveSingleSheetFirstDataRow()
        };
        return true;
    }

    private int AppendImportRowAtEnd(string nameValue)
    {
        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            var nameColumnIndex = GetActiveNameColumnIndex();
            var nameRow = -1;
            foreach (var sheetCfg in _multiSheetSession.Sheets)
            {
                var sheetHeaders = HeaderColumns
                    .Where(h => string.Equals(h.SheetName, sheetCfg.SheetName, StringComparison.Ordinal))
                    .Select(h => new HeaderCell(h.ColumnIndex, h.Name))
                    .ToList();
                var isNameSheet = string.Equals(sheetCfg.SheetName, _multiSheetSession.NameSheetName, StringComparison.Ordinal);
                var addedRow = _workbookService.AddRecordAtSheetLogicalEnd(
                    sheetCfg.SheetName,
                    sheetCfg.FirstDataRow,
                    isNameSheet ? nameColumnIndex : 0,
                    isNameSheet ? nameValue : string.Empty,
                    sheetHeaders);
                if (isNameSheet)
                {
                    nameRow = addedRow;
                }
            }

            return nameRow;
        }

        return _workbookService.AddRecordAtLogicalEnd(
            HeaderRowNumber,
            NameColumnIndex,
            nameValue,
            HeaderColumns.Select(h => new HeaderCell(h.ColumnIndex, h.Name)).ToList());
    }

    public string? RunDataLink(DataLinkProfile profile, out DataLinkResult? result)
    {
        result = null;
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel dữ liệu.";
        }

        if (HeaderColumns.Count == 0)
        {
            return "Chưa có tiêu đề cột trên file gốc.";
        }

        if (string.IsNullOrWhiteSpace(profile.DataFilePath) || !File.Exists(profile.DataFilePath))
        {
            return "File số liệu không tồn tại.";
        }

        if (string.IsNullOrWhiteSpace(profile.DataSheetName))
        {
            return "Chưa chọn sheet file số liệu.";
        }

        if (profile.MainKeyColumnIndex <= 0 || profile.DataKeyColumnIndex <= 0)
        {
            return "Chưa chọn cột key.";
        }

        if (profile.ColumnMappings.Count == 0)
        {
            return "Chưa có cặp cột ánh xạ nào.";
        }

        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);

        if (!TryBuildDataLinkMainContext(profile, out var mainContext, out var contextError))
        {
            return contextError;
        }

        var dataFirstDataRow = Math.Max(profile.DataHeaderRow + 1, 1);
        var formatReferenceRow = SampleRowThreshold > 0 ? SampleRowThreshold : Math.Max(1, HeaderRowNumber);

        try
        {
            using var reader = new ExcelLookupReader();
            reader.Open(profile.DataFilePath, profile.DataSheetName);
            result = _dataLinkService.RunLink(
                _workbookService,
                reader,
                mainContext,
                dataFirstDataRow,
                profile.MainKeyColumnIndex,
                profile.DataKeyColumnIndex,
                profile.ColumnMappings,
                formatReferenceRow);

            SaveDataLinkProfile(profile);

            if (IsAutoSaveEnabled)
            {
                if (!_workbookService.TrySave(out var err))
                {
                    WarnSaveFailed(err);
                    StatusMessage = err ?? "Đã liên kết nhưng chưa ghi được file.";
                }
                else
                {
                    StatusMessage = BuildDataLinkStatusMessage(result);
                }
            }
            else
            {
                _workbookService.CommitPendingEditsToWorksheetOnly();
                StatusMessage = BuildDataLinkStatusMessage(result) + " (chưa ghi file).";
            }

            var selectedLogical = SelectedRecord?.LogicalIndex;
            var selectedRow = SelectedRecord?.RowIndex;
            LoadRecords();
            if (IsMultiSheetMode && selectedLogical is int logicalIdx)
            {
                SelectedRecord = Records.FirstOrDefault(r => r.LogicalIndex == logicalIdx) ?? Records.FirstOrDefault();
            }
            else if (selectedRow is not null)
            {
                SelectedRecord = Records.FirstOrDefault(r => r.RowIndex == selectedRow) ?? Records.FirstOrDefault();
            }
            else
            {
                SelectedRecord = Records.FirstOrDefault();
            }

            if (SelectedRecord is not null)
            {
                LoadEditableFieldsForSelectedRecord();
            }

            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể liên kết số liệu: {root.Message}";
        }
    }

    public string? ExportKeyIssuesReport(DataLinkProfile profile, string outputPath, DataLinkResult? linkResult = null)
    {
        if (!_isWorkbookLoaded)
        {
            return "Chưa mở file Excel dữ liệu.";
        }

        if (string.IsNullOrWhiteSpace(profile.DataFilePath) || !File.Exists(profile.DataFilePath))
        {
            return "File số liệu không tồn tại.";
        }

        if (string.IsNullOrWhiteSpace(profile.DataSheetName))
        {
            return "Chưa chọn sheet file số liệu.";
        }

        if (profile.MainKeyColumnIndex <= 0 || profile.DataKeyColumnIndex <= 0)
        {
            return "Chưa chọn cột key.";
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            return "Đường dẫn file xuất không hợp lệ.";
        }

        if (!TryBuildDataLinkMainContext(profile, out var mainContext, out var contextError))
        {
            return contextError;
        }

        var dataFirstDataRow = Math.Max(profile.DataHeaderRow + 1, 1);

        try
        {
            using var reader = new ExcelLookupReader();
            reader.Open(profile.DataFilePath, profile.DataSheetName);
            var analysis = _dataLinkService.AnalyzeKeyIssues(
                _workbookService,
                reader,
                mainContext,
                dataFirstDataRow,
                profile.MainKeyColumnIndex,
                profile.DataKeyColumnIndex);

            _dataLinkService.ExportReport(outputPath, analysis, linkResult);

            StatusMessage = $"Đã xuất báo cáo key: {outputPath}";
            return null;
        }
        catch (Exception ex)
        {
            var root = GetInnermostException(ex);
            return $"Không thể xuất báo cáo: {root.Message}";
        }
    }

    public DataLinkAnalysis? AnalyzeDataLinkKeys(DataLinkProfile profile)
    {
        if (!_isWorkbookLoaded
            || string.IsNullOrWhiteSpace(profile.DataFilePath)
            || !File.Exists(profile.DataFilePath)
            || string.IsNullOrWhiteSpace(profile.DataSheetName)
            || profile.MainKeyColumnIndex <= 0
            || profile.DataKeyColumnIndex <= 0)
        {
            return null;
        }

        if (!TryBuildDataLinkMainContext(profile, out var mainContext, out _))
        {
            return null;
        }

        var dataFirstDataRow = Math.Max(profile.DataHeaderRow + 1, 1);

        using var reader = new ExcelLookupReader();
        reader.Open(profile.DataFilePath, profile.DataSheetName);
        return _dataLinkService.AnalyzeKeyIssues(
            _workbookService,
            reader,
            mainContext,
            dataFirstDataRow,
            profile.MainKeyColumnIndex,
            profile.DataKeyColumnIndex);
    }

    public HeaderDefinition? GetDefaultDataLinkMainKeyHeader()
    {
        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            var nameColumn = GetActiveNameColumnIndex();
            return HeaderColumns.FirstOrDefault(h =>
                string.Equals(h.SheetName, _multiSheetSession.NameSheetName, StringComparison.Ordinal)
                && h.ColumnIndex == nameColumn);
        }

        return HeaderColumns.FirstOrDefault(h => h.ColumnIndex == NameColumnIndex)
            ?? HeaderColumns.FirstOrDefault();
    }

    private bool TryBuildDataLinkMainContext(
        DataLinkProfile profile,
        out DataLinkMainContext mainContext,
        out string? errorMessage)
    {
        mainContext = new DataLinkMainContext();
        errorMessage = null;

        if (IsMultiSheetMode && _multiSheetSession is not null)
        {
            var keySheetName = string.IsNullOrWhiteSpace(profile.MainKeySheetName)
                ? _multiSheetSession.NameSheetName
                : profile.MainKeySheetName;
            var keySheetCfg = _multiSheetSession.GetSheetConfig(keySheetName);
            if (keySheetCfg is null)
            {
                errorMessage = $"Không tìm thấy cấu hình sheet key '{keySheetName}'.";
                return false;
            }

            var firstDataBySheet = _multiSheetSession.Sheets.ToDictionary(
                s => s.SheetName,
                s => s.FirstDataRow,
                StringComparer.Ordinal);
            var formatRefBySheet = _multiSheetSession.Sheets.ToDictionary(
                s => s.SheetName,
                s => GetSampleRowForSheet(s.SheetName),
                StringComparer.Ordinal);

            mainContext = new DataLinkMainContext
            {
                KeySheetName = keySheetName,
                KeySheetFirstDataRow = keySheetCfg.FirstDataRow,
                FirstDataRowBySheet = firstDataBySheet,
                FormatReferenceRowBySheet = formatRefBySheet
            };
            return true;
        }

        mainContext = new DataLinkMainContext
        {
            KeySheetName = string.Empty,
            KeySheetFirstDataRow = ResolveSingleSheetFirstDataRow()
        };
        return true;
    }

    public string BuildDataLinkSummary(DataLinkResult result)
    {
        var lines = new List<string>();
        if (result.RowsCreatedInMain > 0)
        {
            lines.Add($"Đã tạo {result.RowsCreatedInMain} dòng mới trên file gốc.");
        }

        lines.Add($"Đã khớp {result.RowsMatched} dòng, ghi {result.CellsWritten} ô.");
        lines.Add(BuildSkippedRowsSummary(result.SkippedRows));
        lines.Add(BuildKeyIssuesSummary(
            result.MainFileDuplicates,
            result.DataFileDuplicates,
            result.KeysOnlyInMain,
            result.KeysOnlyInData));

        if (result.SkippedRows.Count > 0)
        {
            lines.Add("Các dòng bị bỏ qua vẫn được ghi chú trong báo cáo chi tiết.");
        }

        return string.Join("\n", lines.Where(line => !string.IsNullOrWhiteSpace(line)));
    }

    public string BuildKeyIssuesSummary(
        IReadOnlyList<DuplicateKeyEntry> mainDuplicates,
        IReadOnlyList<DuplicateKeyEntry> dataDuplicates,
        IReadOnlyList<string> keysOnlyInMain,
        IReadOnlyList<string> keysOnlyInData)
    {
        return $"Key trùng file gốc: {mainDuplicates.Count}\n" +
               $"Key trùng file số liệu: {dataDuplicates.Count}\n" +
               $"Key chỉ có ở file gốc: {keysOnlyInMain.Count}\n" +
               $"Key chỉ có ở file số liệu: {keysOnlyInData.Count}";
    }

    private static string BuildDataLinkStatusMessage(DataLinkResult result)
    {
        var skippedCount = result.SkippedRows.Count;
        if (skippedCount == 0)
        {
            return $"Đã liên kết {result.RowsMatched} dòng, ghi {result.CellsWritten} ô.";
        }

        return $"Đã liên kết {result.RowsMatched} dòng, ghi {result.CellsWritten} ô; bỏ qua {skippedCount} dòng có vấn đề key.";
    }

    private static string BuildSkippedRowsSummary(IReadOnlyList<SkippedLinkEntry> skippedRows)
    {
        if (skippedRows.Count == 0)
        {
            return string.Empty;
        }

        var mainDuplicateCount = skippedRows.Count(r => r.Reason == SkippedLinkReason.MainDuplicateKey);
        var dataDuplicateCount = skippedRows.Count(r => r.Reason == SkippedLinkReason.DataDuplicateKey);
        var noMatchCount = skippedRows.Count(r => r.Reason == SkippedLinkReason.NoMatchInDataFile);

        return $"Bỏ qua {skippedRows.Count} dòng:\n" +
               $"- Key trùng file gốc: {mainDuplicateCount}\n" +
               $"- Key trùng file số liệu: {dataDuplicateCount}\n" +
               $"- Không có số liệu khớp: {noMatchCount}";
    }

    private void ToggleAllHeaders()
    {
        _isBulkUpdatingHeaderVisibility = true;
        try
        {
            if (!_isShowAllTemporarilyEnabled)
            {
                _hiddenHeadersBeforeShowAll = HeaderColumns
                    .Where(h => !h.IsVisible)
                    .Select(h => h.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (var header in HeaderColumns)
                {
                    header.IsVisible = true;
                }

                _isShowAllTemporarilyEnabled = true;
                StatusMessage = "Đang bật hiện tất cả tạm thời.";
            }
            else
            {
                foreach (var header in HeaderColumns)
                {
                    header.IsVisible = !_hiddenHeadersBeforeShowAll.Contains(header.Name);
                }

                _isShowAllTemporarilyEnabled = false;
                StatusMessage = "Đã tắt hiện tất cả, khôi phục trạng thái ẩn đã lưu.";
            }
        }
        finally
        {
            _isBulkUpdatingHeaderVisibility = false;
        }

        LoadEditableFieldsForSelectedRecord();
        OnPropertyChanged(nameof(AreAllHeadersVisible));
        OnPropertyChanged(nameof(ToggleAllHeadersLabel));
    }

    private void ApplyProfileIfAny()
    {
        if (string.IsNullOrWhiteSpace(_profileKey))
        {
            return;
        }

        var hiddenHeaders = _profileStore.TryLoadHiddenHeaders(_profileKey);
        if (hiddenHeaders is null)
        {
            return;
        }

        foreach (var header in HeaderColumns)
        {
            header.IsVisible = !hiddenHeaders.Contains(header.Name);
        }
    }

    private void HeaderOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(HeaderDefinition.IsVisible))
        {
            return;
        }

        if (_isBulkUpdatingHeaderVisibility)
        {
            return;
        }

        if (_isShowAllTemporarilyEnabled)
        {
            _hiddenHeadersBeforeShowAll = HeaderColumns
                .Where(h => !h.IsVisible)
                .Select(h => h.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        if (SelectedRecord is not null)
        {
            LoadEditableFieldsForSelectedRecord();
        }

        RememberHiddenHeaders();
        OnPropertyChanged(nameof(AreAllHeadersVisible));
        OnPropertyChanged(nameof(ToggleAllHeadersLabel));
    }

    private void RefreshOnHeaderDefinitionChanged()
    {
        if (IsMultiSheetMode || !_isWorkbookLoaded || _isRealtimeRefreshInProgress || string.IsNullOrWhiteSpace(SelectedSheet))
        {
            return;
        }

        ApplyHeaderRow();
        LoadRecords();
    }

    private void RefreshOnNameColumnChanged()
    {
        if (IsMultiSheetMode || !_isWorkbookLoaded || _isRealtimeRefreshInProgress || string.IsNullOrWhiteSpace(SelectedSheet))
        {
            return;
        }

        LoadRecords();
    }

    private void RefreshSuggestionFieldsForColumn(string sheetName, int columnIndex)
    {
        if (!AreSuggestionsEnabledForCurrentMode())
        {
            foreach (var field in EditableFields.Where(f => MatchesSuggestionField(f, sheetName, columnIndex)))
            {
                field.SuggestionOptions = [];
            }

            return;
        }

        var sampleRow = string.IsNullOrEmpty(sheetName)
            ? SampleRowThreshold
            : GetSampleRowForSheet(sheetName);
        var hasSuggestions = string.IsNullOrEmpty(sheetName)
            ? _workbookService.ColumnHasSuggestionData(columnIndex, sampleRow)
            : _workbookService.ColumnHasSuggestionData(sheetName, columnIndex, sampleRow);
        var suggestions = hasSuggestions
            ? string.IsNullOrEmpty(sheetName)
                ? _workbookService.GetColumnSuggestions(columnIndex, sampleRow)
                : _workbookService.GetColumnSuggestions(sheetName, columnIndex, sampleRow)
            : [];

        foreach (var field in EditableFields.Where(f => MatchesSuggestionField(f, sheetName, columnIndex)))
        {
            field.SuggestionOptions = suggestions;
        }
    }

    private void RefreshSuggestionFieldsForColumn(int columnIndex) =>
        RefreshSuggestionFieldsForColumn(string.Empty, columnIndex);

    private bool MatchesSuggestionField(EditableField field, string sheetName, int columnIndex)
    {
        if (field.HasDropdown || field.ColumnIndex != columnIndex)
        {
            return false;
        }

        return IsMultiSheetMode
            ? string.Equals(field.SheetName, sheetName, StringComparison.Ordinal)
            : string.IsNullOrEmpty(field.SheetName) || string.Equals(field.SheetName, _workbookService.ActiveSheetName, StringComparison.Ordinal);
    }

    private string ShowInputDialog(string prompt, string defaultValue, string title)
    {
        var dialog = new InputDialog(prompt, defaultValue, title);
        return dialog.ShowDialog() == true ? dialog.InputValue.Trim() : string.Empty;
    }

    public void Dispose()
    {
        PersistSingleSheetSettingsAndBackup();
        FlushPendingChanges(saveToDisk: IsAutoSaveEnabled);
        _workbookService.Dispose();
    }

    private static string BuildDropdownColorGroupKey(string sheetName, string headerName) =>
        $"{sheetName}|{ExtractParentHeaderKey(headerName)}";

    private static string ExtractParentHeaderKey(string headerName)
    {
        if (string.IsNullOrWhiteSpace(headerName))
        {
            return string.Empty;
        }

        var newlineIndex = headerName.IndexOf('\n');
        return newlineIndex < 0
            ? headerName.Trim()
            : headerName[..newlineIndex].Trim();
    }

    private static string ExtractChildHeaderName(string headerName)
    {
        if (string.IsNullOrWhiteSpace(headerName))
        {
            return string.Empty;
        }

        var newlineIndex = headerName.IndexOf('\n');
        return newlineIndex < 0
            ? headerName.Trim()
            : headerName[(newlineIndex + 1)..].Trim();
    }

    private static string NormalizeUiScalePreset(string? preset)
    {
        return preset switch
        {
            "Nhỏ" => "Nhỏ",
            "To" => "To",
            "Rất to" => "Rất to",
            _ => "Vừa"
        };
    }

    private static string SanitizeFileName(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var result = text.Trim();
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            result = result.Replace(c, '_');
        }

        return result;
    }

    private static string BuildUniqueOutputPath(string outputFolder, string baseName)
    {
        var candidate = Path.Combine(outputFolder, $"{baseName}.docx");
        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var index = 2;
        while (true)
        {
            candidate = Path.Combine(outputFolder, $"{baseName}_{index}.docx");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            index++;
        }
    }
}
