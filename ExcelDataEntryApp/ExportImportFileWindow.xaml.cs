using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;
using Microsoft.Win32;

namespace ExcelDataEntryApp;

public partial class ExportImportFileWindow : Window
{
    private readonly MainViewModel _viewModel;
    private bool _isLoadingGenderColumn;

    public ExportImportFileWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        InitializeGenderColumnCombo();
        LoadPreview();
        UpdateSkipColumnsButtonLabel();
        OutputPathTextBox.Text = SuggestOutputPath();
    }

    private void InitializeGenderColumnCombo()
    {
        _isLoadingGenderColumn = true;
        try
        {
            GenderColumnComboBox.ItemsSource = _viewModel.HeaderColumns;
            HeaderDefinition? selected = null;
            if (_viewModel.ImportFileGenderColumnIndex > 0)
            {
                selected = _viewModel.HeaderColumns.FirstOrDefault(h =>
                    h.ColumnIndex == _viewModel.ImportFileGenderColumnIndex);
            }

            selected ??= _viewModel.TryGuessHsskGenderHeader();
            GenderColumnComboBox.SelectedItem = selected;
        }
        finally
        {
            _isLoadingGenderColumn = false;
        }
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
        _viewModel.ImportFileGenderColumnIndex = columnIndex;
        LoadPreview();
    }

    private void UpdateSkipColumnsButtonLabel()
    {
        var count = _viewModel.GetImportFileExportProfile().SkipSampleColumnIndexes.Count;
        SkipColumnsButton.Content = count > 0 ? $"Cột $ Nam ({count})" : "Cột $ Nam";
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        if (!_viewModel.PersistImportFileExportProfileToAppFolder())
        {
            MessageBox.Show(
                "Chưa lưu được — cần mở file Excel và áp dụng tiêu đề trước.",
                "Thiếu thông tin",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show(
            $"Đã lưu profile xuất file import.\n\n"
            + $"File profile:\n{_viewModel.ImportFileExportProfilesFilePath}\n\n"
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
            vm => vm.GetImportFileColumnOptionsForExport(),
            (vm, indexes) => vm.SaveImportFileSkipSampleColumns(indexes),
            vm => vm.PersistImportFileExportProfileToAppFolder(),
            vm => vm.ImportFileExportProfilesFilePath)
        {
            Owner = this
        };
        dialog.ShowDialog();
        UpdateSkipColumnsButtonLabel();
        LoadPreview();
    }

    private void LoadPreview()
    {
        var preview = _viewModel.TryGetImportFileExportPreview();
        if (preview is null)
        {
            PreviewTextBlock.Text = "Chưa xác định được phạm vi xuất. Kiểm tra cấu hình tiêu đề, cột tên và dữ liệu.";
            RootGrid.IsEnabled = false;
            return;
        }

        var profile = _viewModel.GetImportFileExportProfile();
        var skipCount = profile.SkipSampleColumnIndexes.Count;
        var skipLine = skipCount > 0
            ? $"\nCột $ không lấy dòng mẫu (Nam): {skipCount}"
            : string.Empty;
        var genderLine = profile.GenderColumnIndex > 0
            ? $"\nCột giới tính: ${profile.GenderColumnIndex}"
            : "\nCột giới tính: chưa chọn (cột $ Nam áp dụng cho mọi dòng)";

        if (preview.IsMultiSheetMode)
        {
            PreviewTextBlock.Text =
                $"Chế độ nhiều sheet: {preview.SheetCount} sheet.\n"
                + $"Sheet tên: {preview.NameSheetName}\n"
                + $"Dòng dữ liệu: {preview.FirstDataRow} → {preview.LastNameRow} ({preview.RowCount} dòng)\n"
                + $"Hàng mẫu (sheet tên): {preview.SampleRow}"
                + genderLine
                + skipLine;
            return;
        }

        PreviewTextBlock.Text =
            $"Sheet: {preview.NameSheetName}\n"
            + $"Dòng dữ liệu: {preview.FirstDataRow} → {preview.LastNameRow} ({preview.RowCount} dòng)\n"
            + $"Hàng mẫu: {preview.SampleRow}"
            + genderLine
            + skipLine;
    }

    private string SuggestOutputPath()
    {
        var source = _viewModel.FilePath;
        if (string.IsNullOrWhiteSpace(source))
        {
            return string.Empty;
        }

        var directory = Path.GetDirectoryName(source) ?? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var baseName = Path.GetFileNameWithoutExtension(source);
        return Path.Combine(directory, $"{baseName}_import.xlsx");
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Chọn file Excel xuất import",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = Path.GetFileName(OutputPathTextBox.Text.Trim()),
            InitialDirectory = Path.GetDirectoryName(OutputPathTextBox.Text.Trim())
                ?? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        };

        if (dialog.ShowDialog() == true)
        {
            OutputPathTextBox.Text = dialog.FileName;
        }
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        var outputPath = OutputPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            MessageBox.Show("Hãy chọn đường dẫn file xuất.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            outputPath += ".xlsx";
            OutputPathTextBox.Text = outputPath;
        }

        var error = _viewModel.ExportImportFile(outputPath);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Xuất thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show($"Đã xuất file import:\n{outputPath}", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
