using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;
using Microsoft.Win32;

namespace ExcelDataEntryApp;

public partial class ExportHsskWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ObservableCollection<ExportRecordOption> _records;
    private readonly ICollectionView _recordsView;
    private bool _isLoadingGenderColumn;

    public ExportHsskWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        _records = new ObservableCollection<ExportRecordOption>(
            _viewModel.Records.Select(r => new ExportRecordOption
            {
                RowIndex = r.RowIndex,
                Name = r.KeyDisplay
            }));

        _recordsView = CollectionViewSource.GetDefaultView(_records);
        _recordsView.Filter = FilterRecord;
        RecordsListBox.ItemsSource = _recordsView;
        OutputFolderTextBox.Text = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "HSSK_Export");

        InitializeGenderColumnCombo();
        UpdateSkipColumnsButtonLabel();
    }

    private void InitializeGenderColumnCombo()
    {
        _isLoadingGenderColumn = true;
        try
        {
            GenderColumnComboBox.ItemsSource = _viewModel.HeaderColumns;
            HeaderDefinition? selected = null;
            if (_viewModel.HsskGenderColumnIndex > 0)
            {
                selected = _viewModel.HeaderColumns.FirstOrDefault(h => h.ColumnIndex == _viewModel.HsskGenderColumnIndex);
            }

            selected ??= _viewModel.TryGuessHsskGenderHeader();
            GenderColumnComboBox.SelectedItem = selected;
        }
        finally
        {
            _isLoadingGenderColumn = false;
        }
    }

    private void UpdateSkipColumnsButtonLabel()
    {
        var count = _viewModel.GetHsskExportProfile().SkipSampleForMaleColumnIndexes.Count;
        SkipColumnsButton.Content = count > 0 ? $"Cột $ Nam ({count})" : "Cột $ Nam";
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        if (!_viewModel.PersistHsskExportProfileToAppFolder())
        {
            MessageBox.Show(
                "Chưa lưu được — cần mở file Excel và áp dụng tiêu đề trước.",
                "Thiếu thông tin",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show(
            $"Đã lưu profile xuất HSSK.\n\n"
            + $"File profile:\n{_viewModel.HsskExportProfilesFilePath}\n\n"
            + $"File backup:\n{_viewModel.AppConfigBackupFilePath}",
            "Hoàn tất",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void OpenSkipColumns_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SkipSampleColumnsWindow(
            _viewModel,
            "Cột $ không lấy dòng mẫu (Nam)",
            "Chọn cột $: với Nam (hoặc giới tính trống) không lấy dòng mẫu khi ô trống. Nữ vẫn lấy dòng mẫu. Bấm Lưu profile để ghi file cấu hình cạnh thư mục app.",
            vm => vm.GetHsskColumnOptionsForExport(),
            (vm, indexes) => vm.SaveHsskSkipSampleForMaleColumns(indexes),
            vm => vm.PersistHsskExportProfileToAppFolder(),
            vm => vm.HsskExportProfilesFilePath)
        {
            Owner = this
        };
        dialog.ShowDialog();
        UpdateSkipColumnsButtonLabel();
    }

    private bool FilterRecord(object obj)
    {
        if (obj is not ExportRecordOption item)
        {
            return false;
        }

        var q = SearchTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(q))
        {
            return true;
        }

        return item.Name.Contains(q, StringComparison.OrdinalIgnoreCase)
               || item.RowIndex.ToString().Contains(q, StringComparison.OrdinalIgnoreCase);
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn mẫu Word",
            Filter = "Word Document (*.docx)|*.docx",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog() == true)
        {
            TemplatePathTextBox.Text = dialog.FileName;
        }
    }

    private void SetDefaultOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        OutputFolderTextBox.Text = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "HSSK_Export");
    }

    private void SearchTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        _recordsView.Refresh();
    }

    private void GenderColumnComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingGenderColumn)
        {
            return;
        }

        var columnIndex = GenderColumnComboBox.SelectedItem is HeaderDefinition header
            ? header.ColumnIndex
            : 0;
        _viewModel.HsskGenderColumnIndex = columnIndex;
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _recordsView.Cast<ExportRecordOption>())
        {
            item.IsSelected = true;
        }
    }

    private void UnselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _records)
        {
            item.IsSelected = false;
        }
    }

    private void ExportHeadersToExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Lưu file mapping tiêu đề",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "mapping_tieude_placeholder.xlsx"
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var error = _viewModel.ExportHeaderPlaceholdersToExcel(dialog.FileName);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Xuất thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show("Đã xuất file mapping tiêu đề thành công.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        var templatePath = TemplatePathTextBox.Text.Trim();
        var outputFolder = OutputFolderTextBox.Text.Trim();
        var selectedRows = _records.Where(r => r.IsSelected).Select(r => r.RowIndex).ToList();
        var fileNote = FileNoteTextBox.Text.Trim();
        var exportToSingleFile = ExportToSingleFileCheckBox.IsChecked == true;
        var combinedFileName = CombinedFileNameTextBox.Text.Trim();
        var hsskProfile = _viewModel.GetHsskExportProfile();
        var genderColumnIndex = GenderColumnComboBox.SelectedItem is HeaderDefinition genderHeader
            ? genderHeader.ColumnIndex
            : hsskProfile.GenderColumnIndex;
        var femaleOnlyColumns = hsskProfile.SkipSampleForMaleColumnIndexes;

        if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
        {
            MessageBox.Show("Vui lòng chọn file mẫu Word hợp lệ.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (selectedRows.Count == 0)
        {
            MessageBox.Show("Vui lòng chọn ít nhất một dòng để xuất.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        RootGrid.IsEnabled = false;
        ExportProgressBar.Value = 0;
        ExportProgressBar.Visibility = Visibility.Visible;
        ExportProgressTextBlock.Text = "Tiến độ: 0%";
        ExportProgressTextBlock.Visibility = Visibility.Visible;

        IProgress<(int Done, int Total)> progress = new Progress<(int Done, int Total)>(p =>
        {
            var percent = p.Total <= 0 ? 0 : (int)Math.Round((double)p.Done * 100 / p.Total);
            ExportProgressBar.Value = percent;
            ExportProgressTextBlock.Text = $"Tiến độ: {percent}% ({p.Done}/{p.Total})";
        });

        string? error;
        try
        {
            error = await Task.Run(() => _viewModel.ExportHssk(
                templatePath,
                outputFolder,
                selectedRows,
                fileNote,
                exportToSingleFile,
                combinedFileName,
                genderColumnIndex,
                femaleOnlyColumns,
                (done, total) => progress.Report((done, total))));
        }
        finally
        {
            RootGrid.IsEnabled = true;
        }

        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Xuất thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show("Xuất HSSK thành công.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        DialogResult = true;
        Close();
    }
}
