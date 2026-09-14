using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelDataEntryApp.Infrastructure;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;
using Microsoft.Win32;

namespace ExcelDataEntryApp;

public partial class ImportListLinkWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ObservableCollection<HeaderDefinition> _dataHeaders = [];
    private readonly ObservableCollection<ColumnMappingRowItem> _mappings = [];
    private readonly ObservableCollection<string> _savedProfileNames = [];
    private readonly ObservableCollection<ImportListRowOption> _nameOptions = [];
    private readonly Dictionary<int, ImportListColumnFilterState> _columnFilterStates = [];
    private readonly ICollectionView _nameOptionsView;
    private bool _isLoadingProfile;
    private bool _isLoadingSavedProfile;
    private string _activeSearchQuery = string.Empty;

    public ImportListLinkWindow(MainViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        DialogWindowHelper.ClampToWorkArea(this);
        DataContext = this;
        MainHeaders = viewModel.HeaderColumns;
        DataHeaders = _dataHeaders;
        MainFileLabel = BuildFileLabel(viewModel.FilePath, "Chưa có file gốc", "File gốc");
        BackupPathTextBlock.Text = $"Backup tự động cấu hình app: {viewModel.ConfigBackupFilePath}";
        UpdateDataFileLabel();
        MappingsListBox.ItemsSource = _mappings;
        SavedProfileComboBox.ItemsSource = _savedProfileNames;
        _nameOptionsView = CollectionViewSource.GetDefaultView(_nameOptions);
        _nameOptionsView.Filter = FilterNameOption;
        ImportDataGrid.ItemsSource = _nameOptionsView;
        ReloadSavedProfileNames();
        InitializeProfiles();
        SelectNamesTab.Loaded += (_, _) => RefreshNameList(silent: true);
    }

    public ObservableCollection<HeaderDefinition> MainHeaders { get; }
    public ObservableCollection<HeaderDefinition> DataHeaders { get; }
    public string MainFileLabel { get; }

    private void ReloadSavedProfileNames()
    {
        _savedProfileNames.Clear();
        foreach (var name in _viewModel.GetSavedImportListProfileNames())
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

            var namedProfile = _viewModel.TryLoadNamedImportListProfile(lastNamed);
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
        var profile = _viewModel.TryLoadImportListProfile();
        if (profile is not null)
        {
            ApplyProfile(profile);
        }
    }

    private void ApplyProfile(ImportListProfile profile)
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

            DataNameColumnComboBox.SelectedItem = FindHeader(_dataHeaders, profile.DataNameColumnIndex)
                ?? _dataHeaders.FirstOrDefault(h => h.ColumnIndex == profile.DataNameColumnIndex);

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

    private void NewProfile_Click(object sender, RoutedEventArgs e)
    {
        _isLoadingSavedProfile = true;
        SavedProfileComboBox.SelectedItem = null;
        _isLoadingSavedProfile = false;

        DataFilePathTextBox.Text = string.Empty;
        UpdateDataFileLabel();
        DataSheetComboBox.ItemsSource = null;
        DataSheetComboBox.SelectedItem = null;
        DataHeaderRowTextBox.Text = "1";
        _dataHeaders.Clear();
        DataNameColumnComboBox.ItemsSource = null;
        DataNameColumnComboBox.SelectedItem = null;
        _mappings.Clear();
        SyncColumnFilterStatesAndGrid();
        _nameOptions.Clear();
        _nameOptionsView.Refresh();
    }

    private void ClearDataFile_Click(object sender, RoutedEventArgs e) => ClearDataFileState();

    private void ClearDataFileState()
    {
        DataFilePathTextBox.Text = string.Empty;
        UpdateDataFileLabel();
        DataSheetComboBox.ItemsSource = null;
        DataSheetComboBox.SelectedItem = null;
        _dataHeaders.Clear();
        DataNameColumnComboBox.ItemsSource = null;
        DataNameColumnComboBox.SelectedItem = null;

        foreach (var mapping in _mappings)
        {
            mapping.SourceColumn = null;
        }

        SyncColumnFilterStatesAndGrid();
        _nameOptions.Clear();
        _nameOptionsView.Refresh();
    }

    private void BrowseDataFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx",
            Title = "Chọn file danh sách"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        DataFilePathTextBox.Text = dialog.FileName;
        UpdateDataFileLabel();
        LoadSheets(dialog.FileName, null);
        ReloadDataHeaders();
        RefreshNameList(silent: true);
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
            MessageBox.Show($"Không đọc được file danh sách: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void DataSourceSettings_Changed(object sender, RoutedEventArgs e)
    {
        if (_isLoadingProfile)
        {
            return;
        }

        ReloadDataHeaders();
        RefreshNameList(silent: true);
    }

    private void ReloadDataHeaders()
    {
        _dataHeaders.Clear();
        DataNameColumnComboBox.ItemsSource = null;

        var path = DataFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            SyncColumnFilterStatesAndGrid();
            return;
        }

        if (DataSheetComboBox.SelectedItem is not string sheetName)
        {
            SyncColumnFilterStatesAndGrid();
            return;
        }

        if (!int.TryParse(DataHeaderRowTextBox.Text.Trim(), out var headerRow) || headerRow < 1)
        {
            SyncColumnFilterStatesAndGrid();
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

            DataNameColumnComboBox.ItemsSource = _dataHeaders;
            if (DataNameColumnComboBox.SelectedItem is null)
            {
                DataNameColumnComboBox.SelectedItem = _dataHeaders.FirstOrDefault();
            }

            SyncColumnFilterStatesAndGrid();
            RefreshMappingColumns();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không đọc được tiêu đề file danh sách: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SyncColumnFilterStatesAndGrid()
    {
        var activeColumnIndexes = _dataHeaders.Select(h => h.ColumnIndex).ToHashSet();
        foreach (var columnIndex in _columnFilterStates.Keys.Where(i => !activeColumnIndexes.Contains(i)).ToList())
        {
            _columnFilterStates.Remove(columnIndex);
        }

        foreach (var header in _dataHeaders)
        {
            if (!_columnFilterStates.TryGetValue(header.ColumnIndex, out var state))
            {
                state = new ImportListColumnFilterState
                {
                    ColumnIndex = header.ColumnIndex,
                    HeaderName = header.Name
                };
                _columnFilterStates[header.ColumnIndex] = state;
            }
            else
            {
                state.HeaderName = header.Name;
            }
        }

        BuildImportDataGridColumns();
    }

    private void BuildImportDataGridColumns()
    {
        ImportDataGrid.Columns.Clear();

        ImportDataGrid.Columns.Add(new DataGridCheckBoxColumn
        {
            Header = "Chọn",
            Binding = new Binding(nameof(ImportListRowOption.IsSelected))
            {
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = 56
        });

        foreach (var header in _dataHeaders)
        {
            if (!_columnFilterStates.TryGetValue(header.ColumnIndex, out var filterState))
            {
                continue;
            }

            ImportDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = CreateFilterHeader(header, filterState),
                Binding = new Binding($"[{header.ColumnIndex}]"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                MinWidth = 90,
                IsReadOnly = true
            });
        }
    }

    private FrameworkElement CreateFilterHeader(HeaderDefinition header, ImportListColumnFilterState filterState)
    {
        var grid = new Grid
        {
            Background = Brushes.Transparent,
            ToolTip = header.Name
        };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var title = new TextBlock
        {
            Text = header.Name,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 2, 0)
        };
        Grid.SetColumn(title, 0);

        var filterButton = new Button
        {
            Padding = new Thickness(4, 0, 4, 0),
            MinWidth = 26,
            MinHeight = 22,
            Margin = new Thickness(0, 0, 2, 0),
            ToolTip = "Lọc cột (giống Excel)"
        };
        filterButton.SetBinding(ContentControl.ContentProperty, new Binding(nameof(ImportListColumnFilterState.FilterGlyph))
        {
            Source = filterState
        });
        filterButton.Click += (_, _) =>
        {
            var distinctValues = GetDistinctColumnValues(header.ColumnIndex);
            ImportListColumnFilterPopup.Show(filterButton, filterState, distinctValues, _ =>
            {
                _nameOptionsView.Refresh();
            });
        };
        Grid.SetColumn(filterButton, 1);

        grid.Children.Add(title);
        grid.Children.Add(filterButton);
        return grid;
    }

    private IReadOnlyList<string> GetDistinctColumnValues(int columnIndex) =>
        _nameOptions
            .Select(row => row.GetCellText(columnIndex))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList();

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

        var profile = _viewModel.TryLoadNamedImportListProfile(profileName);
        if (profile is null)
        {
            return;
        }

        _viewModel.RememberLastNamedMappingProfile(profileName);
        ApplyProfile(profile);
        RefreshNameList(silent: true);
    }

    private void LoadProfile_Click(object sender, RoutedEventArgs e)
    {
        if (SavedProfileComboBox.SelectedItem is not string profileName)
        {
            MessageBox.Show("Chọn profile cần tải.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var profile = _viewModel.TryLoadNamedImportListProfile(profileName);
        if (profile is null)
        {
            MessageBox.Show("Không tìm thấy profile đã chọn.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _viewModel.RememberLastNamedMappingProfile(profileName);
        ApplyProfile(profile);
        RefreshNameList(silent: true);
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
        var dialog = new InputDialog("Nhập tên profile:", defaultName, "Lưu profile");
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

        _viewModel.SaveNamedImportListProfile(profileName, profile);
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
            FileName = $"ImportListProfiles_Backup_{DateTime.Now:yyyyMMdd_HHmmss}.json",
            Title = "Backup profile import danh sách"
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
            Title = "Restore profile import danh sách"
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
        if (restoredNames.Count == 0)
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
            var profile = _viewModel.TryLoadNamedImportListProfile(selectedName);
            if (profile is not null)
            {
                _viewModel.RememberLastNamedMappingProfile(selectedName);
                ApplyProfile(profile);
                RefreshNameList(silent: true);
            }
        }

        MessageBox.Show($"Đã restore {restoredNames.Count} profile từ file backup.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BackupAppConfig_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.ExportConfigBackup();
        MessageBox.Show(
            $"Đã lưu backup cấu hình tại:\n{_viewModel.ConfigBackupFilePath}",
            "Backup cấu hình",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void RefreshNameList_Click(object sender, RoutedEventArgs e) => RefreshNameList(silent: false);

    private void RefreshNameList(bool silent)
    {
        if (!TryBuildProfile(out var profile, out var validationError, requireMappings: false))
        {
            if (!silent)
            {
                MessageBox.Show(validationError, "Chưa cấu hình đủ", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            return;
        }

        _nameOptions.Clear();
        foreach (var option in _viewModel.LoadImportListRowOptions(profile))
        {
            _nameOptions.Add(option);
        }

        SyncColumnFilterStatesAndGrid();
        _nameOptionsView.Refresh();

        if (!silent)
        {
            var visible = _nameOptionsView.Cast<ImportListRowOption>().Count();
            var selected = _nameOptions.Count(o => o.IsSelected);
            MessageBox.Show(
                $"Đã tải {_nameOptions.Count} tên ({visible} đang hiển thị sau lọc, {selected} đang chọn).",
                "Danh sách tên",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }

    private bool FilterNameOption(object obj)
    {
        if (obj is not ImportListRowOption item)
        {
            return false;
        }

        foreach (var filterState in _columnFilterStates.Values)
        {
            if (filterState.AllowedValues is null)
            {
                continue;
            }

            if (filterState.AllowedValues.Count == 0)
            {
                return false;
            }

            var cellValue = item.GetCellText(filterState.ColumnIndex);
            if (!filterState.AllowedValues.Contains(cellValue))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(_activeSearchQuery)
            && !VietnameseTextHelper.ContainsNormalized(item.RowSearchText, _activeSearchQuery)
            && !VietnameseTextHelper.ContainsNormalized(item.KeyValue, _activeSearchQuery)
            && !item.DataRowIndex.ToString(CultureInfo.InvariantCulture).Contains(_activeSearchQuery, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private void ApplySearch_Click(object sender, RoutedEventArgs e)
    {
        _activeSearchQuery = NameSearchTextBox.Text?.Trim() ?? string.Empty;
        _nameOptionsView.Refresh();
    }

    private void ClearFilters_Click(object sender, RoutedEventArgs e)
    {
        foreach (var filterState in _columnFilterStates.Values)
        {
            filterState.AllowedValues = null;
        }

        _nameOptionsView.Refresh();
    }

    private void SelectAllNames_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _nameOptionsView.Cast<ImportListRowOption>().Where(o => o.CanSelect))
        {
            item.IsSelected = true;
        }
    }

    private void UnselectAllNames_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _nameOptions)
        {
            item.IsSelected = false;
        }
    }

    private void DataNameColumn_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingProfile)
        {
            return;
        }

        RefreshNameList(silent: true);
    }

    private void RunImport_Click(object sender, RoutedEventArgs e)
    {
        if (!TryBuildProfile(out var profile, out var validationError))
        {
            MessageBox.Show(validationError, "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (_nameOptions.Count == 0)
        {
            RefreshNameList(silent: true);
        }

        var selectedRows = _nameOptions
            .Where(o => o.IsSelected && o.CanSelect)
            .Select(o => o.DataRowIndex)
            .Distinct()
            .ToList();

        if (selectedRows.Count == 0)
        {
            MessageBox.Show(
                "Chưa chọn tên nào. Mở tab \"1. Chọn tên import\" và chọn ít nhất một người.",
                "Thiếu thông tin",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var request = new ImportListLinkRequest
        {
            Profile = profile,
            SelectedDataRowIndices = selectedRows
        };

        var error = _viewModel.RunImportListLink(request, out var result);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Import thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (result is null)
        {
            return;
        }

        var summary = _viewModel.BuildImportListLinkSummary(result);
        MessageBox.Show(summary, "Import hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        DialogResult = true;
        Close();
    }

    private bool TryBuildProfile(out ImportListProfile profile, out string validationError, bool requireMappings = true)
    {
        profile = new ImportListProfile();
        validationError = string.Empty;

        var path = DataFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            validationError = "Vui lòng chọn file danh sách hợp lệ (tab 2).";
            return false;
        }

        if (DataSheetComboBox.SelectedItem is not string sheetName)
        {
            validationError = "Vui lòng chọn sheet file danh sách (tab 2).";
            return false;
        }

        if (!int.TryParse(DataHeaderRowTextBox.Text.Trim(), out var headerRow) || headerRow < 1)
        {
            validationError = "Dòng tiêu đề file danh sách phải là số >= 1.";
            return false;
        }

        if (DataNameColumnComboBox.SelectedItem is not HeaderDefinition nameColumn)
        {
            validationError = "Vui lòng chọn cột tên trên file danh sách (tab 2).";
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
            validationError = "Vui lòng thêm ít nhất một cặp cột ánh xạ hợp lệ (tab 2).";
            return false;
        }

        profile = new ImportListProfile
        {
            DataFilePath = path,
            DataSheetName = sheetName,
            DataHeaderRow = headerRow,
            DataNameColumnIndex = nameColumn.ColumnIndex,
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

    private static HeaderDefinition? FindHeader(IEnumerable<HeaderDefinition> headers, int columnIndex) =>
        headers.FirstOrDefault(h => h.ColumnIndex == columnIndex);

    private static HeaderDefinition CreatePlaceholderHeader(string sheetName, int columnIndex) =>
        new()
        {
            Name = $"Cột {columnIndex}",
            ColumnIndex = columnIndex,
            SheetName = sheetName ?? string.Empty
        };

    private static HeaderDefinition CreatePlaceholderHeader(int columnIndex) =>
        CreatePlaceholderHeader(string.Empty, columnIndex);

    private void UpdateDataFileLabel()
    {
        DataFileLabelTextBlock.Text = BuildFileLabel(
            DataFilePathTextBox.Text.Trim(),
            "Chưa chọn file danh sách",
            "File danh sách");
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
