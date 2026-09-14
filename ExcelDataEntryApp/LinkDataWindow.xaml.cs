using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelDataEntryApp.Infrastructure;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;
using Microsoft.Win32;

namespace ExcelDataEntryApp;

public partial class LinkDataWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ObservableCollection<HeaderDefinition> _dataHeaders = [];
    private readonly ObservableCollection<ColumnMappingRowItem> _mappings = [];
    private readonly ObservableCollection<string> _savedProfileNames = [];
    private bool _isLoadingProfile;
    private bool _isLoadingSavedProfile;

    public LinkDataWindow(MainViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        DialogWindowHelper.ClampToWorkArea(this);
        DataContext = this;
        MainHeaders = viewModel.HeaderColumns;
        DataHeaders = _dataHeaders;
        MainFileLabel = BuildFileLabel(viewModel.FilePath, "Chưa có file gốc", "File gốc");
        UpdateDataFileLabel();
        MainKeyColumnComboBox.ItemsSource = viewModel.HeaderColumns;
        MappingsListBox.ItemsSource = _mappings;
        SavedProfileComboBox.ItemsSource = _savedProfileNames;
        ReloadSavedProfileNames();
        InitializeProfiles();
        UpdateMultiSheetHint();
    }

    private void UpdateMultiSheetHint()
    {
        if (_viewModel.IsMultiSheetMode)
        {
            HintTextBlock.Text =
                "Chế độ nhiều sheet: dùng tiêu đề đã ghép (kèm tên sheet). Key thường là cột tên trên sheet tên; "
                + "cột đích có thể thuộc sheet khác và sẽ ghi theo cùng chỉ số dòng logic. "
                + "Dòng có key trùng hoặc không khớp sẽ bỏ qua; các dòng khác vẫn liên kết bình thường.";
        }
    }

    public ObservableCollection<HeaderDefinition> MainHeaders { get; }
    public ObservableCollection<HeaderDefinition> DataHeaders { get; }
    public string MainFileLabel { get; }

    private void ReloadSavedProfileNames()
    {
        _savedProfileNames.Clear();
        foreach (var name in _viewModel.GetSavedMappingProfileNames())
        {
            _savedProfileNames.Add(name);
        }
    }

    private void InitializeProfiles()
    {
        var lastNamed = _viewModel.TryLoadLastNamedMappingProfileName();
        if (!string.IsNullOrWhiteSpace(lastNamed)
            && _savedProfileNames.Any(n => string.Equals(n, lastNamed, StringComparison.OrdinalIgnoreCase)))
        {
            _isLoadingSavedProfile = true;
            SavedProfileComboBox.SelectedItem = _savedProfileNames.FirstOrDefault(n =>
                string.Equals(n, lastNamed, StringComparison.OrdinalIgnoreCase));
            _isLoadingSavedProfile = false;

            var namedProfile = _viewModel.TryLoadNamedMappingProfile(lastNamed);
            if (namedProfile is not null)
            {
                ApplyProfile(namedProfile);
                return;
            }
        }

        LoadSavedProfile();
    }

    private void LoadSavedProfile()
    {
        var profile = _viewModel.TryLoadDataLinkProfile();
        if (profile is null)
        {
            MainKeyColumnComboBox.SelectedItem = _viewModel.GetDefaultDataLinkMainKeyHeader()
                ?? _viewModel.HeaderColumns.FirstOrDefault();
            return;
        }

        ApplyProfile(profile);
    }

    private void ApplyProfile(DataLinkProfile profile)
    {
        _isLoadingProfile = true;
        try
        {
            DataFilePathTextBox.Text = profile.DataFilePath;
            UpdateDataFileLabel();
            DataHeaderRowTextBox.Text = profile.DataHeaderRow.ToString(CultureInfo.InvariantCulture);

            if (!string.IsNullOrWhiteSpace(profile.DataFilePath) && File.Exists(profile.DataFilePath))
            {
                LoadSheets(profile.DataFilePath, profile.DataSheetName);
                ReloadDataHeaders();
            }
            else if (!string.IsNullOrWhiteSpace(profile.DataSheetName))
            {
                DataSheetComboBox.ItemsSource = new[] { profile.DataSheetName };
                DataSheetComboBox.SelectedItem = profile.DataSheetName;
                ApplyDataHeadersFromProfile(profile);
            }
            else
            {
                ClearDataFileState(clearPath: false);
            }

            MainKeyColumnComboBox.SelectedItem = FindMainHeader(profile.MainKeySheetName, profile.MainKeyColumnIndex)
                ?? _viewModel.GetDefaultDataLinkMainKeyHeader()
                ?? _viewModel.HeaderColumns.FirstOrDefault();
            DataKeyColumnComboBox.SelectedItem = FindHeader(_dataHeaders, profile.DataKeyColumnIndex)
                ?? _dataHeaders.FirstOrDefault(h => h.ColumnIndex == profile.DataKeyColumnIndex);

            _mappings.Clear();
            foreach (var mapping in profile.ColumnMappings)
            {
                _mappings.Add(new ColumnMappingRowItem
                {
                    TargetColumn = FindMainHeader(mapping.TargetSheetName, mapping.TargetColumnIndex)
                        ?? CreatePlaceholderHeader(mapping.TargetSheetName, mapping.TargetColumnIndex),
                    SourceColumn = FindHeader(_dataHeaders, mapping.SourceColumnIndex)
                        ?? CreatePlaceholderHeader(mapping.SourceColumnIndex)
                });
            }
        }
        finally
        {
            _isLoadingProfile = false;
        }
    }

    private void ApplyDataHeadersFromProfile(DataLinkProfile profile)
    {
        _dataHeaders.Clear();
        var columnIndexes = profile.ColumnMappings
            .Select(m => m.SourceColumnIndex)
            .Append(profile.DataKeyColumnIndex)
            .Where(i => i > 0)
            .Distinct()
            .OrderBy(i => i);

        foreach (var columnIndex in columnIndexes)
        {
            _dataHeaders.Add(CreatePlaceholderHeader(columnIndex));
        }

        DataKeyColumnComboBox.ItemsSource = _dataHeaders;
    }

    private void ClearDataFile_Click(object sender, RoutedEventArgs e)
    {
        ClearDataFileState(clearPath: true);
    }

    private void ClearDataFileState(bool clearPath)
    {
        if (clearPath)
        {
            DataFilePathTextBox.Text = string.Empty;
        }

        UpdateDataFileLabel();
        DataSheetComboBox.ItemsSource = null;
        DataSheetComboBox.SelectedItem = null;
        _dataHeaders.Clear();
        DataKeyColumnComboBox.ItemsSource = null;
        DataKeyColumnComboBox.SelectedItem = null;

        foreach (var mapping in _mappings)
        {
            mapping.SourceColumn = null;
        }
    }

    private void NewProfile_Click(object sender, RoutedEventArgs e)
    {
        ResetToNewProfile();
    }

    private void ResetToNewProfile()
    {
        _isLoadingSavedProfile = true;
        SavedProfileComboBox.SelectedItem = null;
        _isLoadingSavedProfile = false;

        ClearDataFileState(clearPath: true);
        DataHeaderRowTextBox.Text = "1";
        _mappings.Clear();

        MainKeyColumnComboBox.SelectedItem = _viewModel.GetDefaultDataLinkMainKeyHeader()
            ?? _viewModel.HeaderColumns.FirstOrDefault();
    }

    private void LoadProfile_Click(object sender, RoutedEventArgs e)
    {
        if (SavedProfileComboBox.SelectedItem is not string profileName)
        {
            MessageBox.Show("Chọn profile cần tải.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var profile = _viewModel.TryLoadNamedMappingProfile(profileName);
        if (profile is null)
        {
            MessageBox.Show("Không tìm thấy profile đã chọn.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _viewModel.RememberLastNamedMappingProfile(profileName);
        ApplyProfile(profile);
    }

    private void BrowseDataFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx",
            Title = "Chọn file số liệu"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        DataFilePathTextBox.Text = dialog.FileName;
        UpdateDataFileLabel();
        LoadSheets(dialog.FileName, null);
        ReloadDataHeaders();
    }

    private void LoadSheets(string path, string? selectedSheet)
    {
        DataSheetComboBox.ItemsSource = null;
        try
        {
            var sheets = _viewModel.GetDataFileSheetNames(path);
            DataSheetComboBox.ItemsSource = sheets;
            if (!string.IsNullOrWhiteSpace(selectedSheet) && sheets.Contains(selectedSheet))
            {
                DataSheetComboBox.SelectedItem = selectedSheet;
            }
            else
            {
                DataSheetComboBox.SelectedIndex = sheets.Count > 0 ? 0 : -1;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không đọc được file số liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void DataSourceSettings_Changed(object sender, RoutedEventArgs e)
    {
        if (_isLoadingProfile)
        {
            return;
        }

        ReloadDataHeaders();
    }

    private void ReloadDataHeaders()
    {
        _dataHeaders.Clear();
        DataKeyColumnComboBox.ItemsSource = null;

        var path = DataFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return;
        }

        if (DataSheetComboBox.SelectedItem is not string sheetName)
        {
            return;
        }

        if (!int.TryParse(DataHeaderRowTextBox.Text.Trim(), out var headerRow) || headerRow < 1)
        {
            return;
        }

        try
        {
            foreach (var header in _viewModel.LoadDataFileHeaders(path, sheetName, headerRow))
            {
                _dataHeaders.Add(new HeaderDefinition
                {
                    Name = header.Name,
                    ColumnIndex = header.ColumnIndex
                });
            }

            DataKeyColumnComboBox.ItemsSource = _dataHeaders;
            if (DataKeyColumnComboBox.SelectedItem is null)
            {
                DataKeyColumnComboBox.SelectedItem = _dataHeaders.FirstOrDefault();
            }

            RefreshMappingColumns();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không đọc được tiêu đề file số liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void RefreshMappingColumns()
    {
        foreach (var mapping in _mappings)
        {
            if (mapping.TargetColumn is not null)
            {
                mapping.TargetColumn = FindMainHeader(mapping.TargetColumn.SheetName, mapping.TargetColumn.ColumnIndex)
                    ?? mapping.TargetColumn;
            }

            if (mapping.SourceColumn is not null)
            {
                mapping.SourceColumn = FindHeader(_dataHeaders, mapping.SourceColumn.ColumnIndex)
                    ?? mapping.SourceColumn;
            }
        }
    }

    private void AddMapping_Click(object sender, RoutedEventArgs e)
    {
        _mappings.Add(new ColumnMappingRowItem
        {
            TargetColumn = _viewModel.HeaderColumns.FirstOrDefault(),
            SourceColumn = _dataHeaders.FirstOrDefault()
        });
        MappingsListBox.ScrollIntoView(_mappings[^1]);
    }

    private void RemoveMapping_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: ColumnMappingRowItem item })
        {
            return;
        }

        _mappings.Remove(item);
    }

    private void SavedProfileComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingSavedProfile || _isLoadingProfile)
        {
            return;
        }

        if (SavedProfileComboBox.SelectedItem is not string profileName)
        {
            return;
        }

        var profile = _viewModel.TryLoadNamedMappingProfile(profileName);
        if (profile is null)
        {
            return;
        }

        _viewModel.RememberLastNamedMappingProfile(profileName);
        ApplyProfile(profile);
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildProfile(out var profile, out var validationError))
        {
            MessageBox.Show(validationError, "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var defaultName = SavedProfileComboBox.SelectedItem as string
            ?? _savedProfileNames.FirstOrDefault()
            ?? "Profile 1";
        var dialog = new InputDialog("Nhập tên profile xếp cặp:", defaultName, "Lưu profile");
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var profileName = dialog.InputValue.Trim();
        if (string.IsNullOrWhiteSpace(profileName))
        {
            MessageBox.Show("Tên profile không được để trống.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _viewModel.SaveNamedMappingProfile(profileName, profile);
        ReloadSavedProfileNames();
        _isLoadingSavedProfile = true;
        SavedProfileComboBox.SelectedItem = _savedProfileNames.FirstOrDefault(n =>
            string.Equals(n, profileName, StringComparison.OrdinalIgnoreCase));
        _isLoadingSavedProfile = false;

        MessageBox.Show($"Đã lưu profile \"{profileName}\".", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void DeleteProfile_Click(object sender, RoutedEventArgs e)
    {
        if (SavedProfileComboBox.SelectedItem is not string profileName)
        {
            MessageBox.Show("Chọn profile cần xóa.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            $"Xóa profile \"{profileName}\"?",
            "Xác nhận",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        _viewModel.DeleteNamedMappingProfile(profileName);
        ReloadSavedProfileNames();
        SavedProfileComboBox.SelectedItem = null;
    }

    private void BackupProfile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "JSON files (*.json)|*.json",
            FileName = $"DataLinkProfiles_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.json",
            Title = "Backup profile liên kết"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var error = _viewModel.BackupNamedMappingProfiles(dialog.FileName);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Backup thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show($"Đã backup profile ra file:\n{dialog.FileName}", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void RestoreProfile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON files (*.json)|*.json",
            Title = "Restore profile liên kết từ file backup"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var error = _viewModel.RestoreNamedMappingProfiles(dialog.FileName, out var restoredNames);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Restore thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ReloadSavedProfileNames();
        var restoredCount = restoredNames.Count;
        if (restoredCount == 0)
        {
            SavedProfileComboBox.SelectedItem = null;
            MessageBox.Show("Restore hoàn tất nhưng không có profile hợp lệ trong file backup.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var lastNamed = _viewModel.TryLoadLastNamedMappingProfileName();
        var selectedName = !string.IsNullOrWhiteSpace(lastNamed)
            ? _savedProfileNames.FirstOrDefault(n => string.Equals(n, lastNamed, StringComparison.OrdinalIgnoreCase))
            : null;
        selectedName ??= _savedProfileNames.FirstOrDefault(n => string.Equals(n, restoredNames[0], StringComparison.OrdinalIgnoreCase))
            ?? _savedProfileNames.FirstOrDefault();

        if (!string.IsNullOrWhiteSpace(selectedName))
        {
            _isLoadingSavedProfile = true;
            SavedProfileComboBox.SelectedItem = selectedName;
            _isLoadingSavedProfile = false;
            var profile = _viewModel.TryLoadNamedMappingProfile(selectedName);
            if (profile is not null)
            {
                _viewModel.RememberLastNamedMappingProfile(selectedName);
                ApplyProfile(profile);
            }
        }

        MessageBox.Show($"Đã restore {restoredCount} profile từ file backup.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void RunLink_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildProfile(out var profile, out var validationError))
        {
            MessageBox.Show(validationError, "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var error = _viewModel.RunDataLink(profile, out var result);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Liên kết thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (result is null)
        {
            return;
        }

        var successSummary = _viewModel.BuildDataLinkSummary(result);
        if (result.HasNotableIssues)
        {
            var exportReport = MessageBox.Show(
                successSummary + "\n\nBạn có muốn xuất báo cáo chi tiết?",
                result.RowsMatched > 0 ? "Liên kết hoàn tất" : "Không có dòng nào được liên kết",
                MessageBoxButton.YesNo,
                result.RowsMatched > 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);

            if (exportReport == MessageBoxResult.Yes)
            {
                ExportKeyIssuesReport(profile, result);
            }
        }
        else
        {
            MessageBox.Show(successSummary, "Liên kết thành công", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void ExportReport_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildProfile(out var profile, out var validationError, requireMappings: false))
        {
            MessageBox.Show(validationError, "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ExportKeyIssuesReport(profile);
    }

    private void ExportKeyIssuesReport(DataLinkProfile profile, DataLinkResult? linkResult = null)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx",
            FileName = $"BaoCao_Key_LienKet_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
            Title = "Xuất báo cáo key liên kết"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var error = _viewModel.ExportKeyIssuesReport(profile, dialog.FileName, linkResult);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Xuất báo cáo thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show($"Đã xuất báo cáo:\n{dialog.FileName}", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private bool TryBuildProfile(out DataLinkProfile profile, out string validationError, bool requireMappings = true)
    {
        profile = new DataLinkProfile();
        validationError = string.Empty;

        var path = DataFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            validationError = "Vui lòng chọn file số liệu hợp lệ.";
            return false;
        }

        if (DataSheetComboBox.SelectedItem is not string sheetName)
        {
            validationError = "Vui lòng chọn sheet file số liệu.";
            return false;
        }

        if (!int.TryParse(DataHeaderRowTextBox.Text.Trim(), out var headerRow) || headerRow < 1)
        {
            validationError = "Dòng tiêu đề file số liệu phải là số >= 1.";
            return false;
        }

        if (MainKeyColumnComboBox.SelectedItem is not HeaderDefinition mainKey)
        {
            validationError = "Vui lòng chọn cột key file gốc.";
            return false;
        }

        if (_viewModel.IsMultiSheetMode && string.IsNullOrWhiteSpace(mainKey.SheetName))
        {
            validationError = "Cột key phải thuộc một sheet khi đang ghép nhiều sheet.";
            return false;
        }

        if (DataKeyColumnComboBox.SelectedItem is not HeaderDefinition dataKey)
        {
            validationError = "Vui lòng chọn cột key file số liệu.";
            return false;
        }

        var validMappings = _mappings
            .Where(m => m.TargetColumn is not null && m.SourceColumn is not null)
            .Select(m => new ColumnMappingPair
            {
                TargetColumnIndex = m.TargetColumn!.ColumnIndex,
                TargetSheetName = m.TargetColumn.SheetName ?? string.Empty,
                SourceColumnIndex = m.SourceColumn!.ColumnIndex
            })
            .ToList();

        if (requireMappings && validMappings.Count == 0)
        {
            validationError = "Vui lòng thêm ít nhất một cặp cột ánh xạ hợp lệ.";
            return false;
        }

        profile = new DataLinkProfile
        {
            DataFilePath = path,
            DataSheetName = sheetName,
            DataHeaderRow = headerRow,
            MainKeyColumnIndex = mainKey.ColumnIndex,
            MainKeySheetName = mainKey.SheetName ?? string.Empty,
            DataKeyColumnIndex = dataKey.ColumnIndex,
            ColumnMappings = validMappings
        };

        return true;
    }

    private HeaderDefinition? FindMainHeader(string? sheetName, int columnIndex)
    {
        if (_viewModel.IsMultiSheetMode)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                return _viewModel.HeaderColumns.FirstOrDefault(h =>
                    h.ColumnIndex == columnIndex
                    && string.Equals(h.SheetName, sheetName, StringComparison.Ordinal));
            }

            return _viewModel.HeaderColumns.FirstOrDefault(h => h.ColumnIndex == columnIndex);
        }

        return _viewModel.HeaderColumns.FirstOrDefault(h =>
            h.ColumnIndex == columnIndex
            && string.IsNullOrEmpty(h.SheetName));
    }

    private static HeaderDefinition? FindHeader(IEnumerable<HeaderDefinition> headers, int columnIndex)
    {
        return headers.FirstOrDefault(h => h.ColumnIndex == columnIndex);
    }

    private static HeaderDefinition CreatePlaceholderHeader(string sheetName, int columnIndex)
    {
        return new HeaderDefinition
        {
            Name = $"Cột {columnIndex}",
            ColumnIndex = columnIndex,
            SheetName = sheetName ?? string.Empty
        };
    }

    private static HeaderDefinition CreatePlaceholderHeader(int columnIndex) =>
        CreatePlaceholderHeader(string.Empty, columnIndex);

    private void UpdateDataFileLabel()
    {
        DataFileLabelTextBlock.Text = BuildFileLabel(
            DataFilePathTextBox.Text.Trim(),
            "Chưa chọn file",
            "File số liệu");
        DataFileLabelTextBlock.ToolTip = string.IsNullOrWhiteSpace(DataFilePathTextBox.Text.Trim())
            ? null
            : DataFilePathTextBox.Text.Trim();
    }

    private static string BuildFileLabel(string filePath, string fallback, string prefix)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return fallback;
        }

        var fileName = Path.GetFileName(filePath);
        return string.IsNullOrWhiteSpace(fileName) ? fallback : $"{prefix}: {fileName}";
    }
}
