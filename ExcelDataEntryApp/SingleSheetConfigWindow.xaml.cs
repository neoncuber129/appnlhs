using System.Globalization;
using System.Windows;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class SingleSheetConfigWindow : Window
{
    private readonly MainViewModel _viewModel;

    public SingleSheetConfigWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        HeaderRowTextBox.Text = _viewModel.HeaderRowNumber.ToString(CultureInfo.InvariantCulture);
        NameColumnTextBox.Text = _viewModel.NameColumnIndex.ToString(CultureInfo.InvariantCulture);
        SampleRowTextBox.Text = _viewModel.SampleRowThreshold.ToString(CultureInfo.InvariantCulture);
        AutoSkipBlankCheckBox.IsChecked = _viewModel.AutoSkipBlankHeaders;
        SuggestionsCheckBox.IsChecked = _viewModel.IsSuggestionsEnabled;
        AutoSaveCheckBox.IsChecked = _viewModel.IsAutoSaveEnabled;
        BackupPathTextBlock.Text = $"Backup tự động: {_viewModel.ConfigBackupFilePath}";
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadPositiveInt(HeaderRowTextBox.Text, out var headerRow, "Hàng tiêu đề"))
        {
            return;
        }

        if (!TryReadPositiveInt(NameColumnTextBox.Text, out var nameColumn, "Cột tên"))
        {
            return;
        }

        if (!TryReadNonNegativeInt(SampleRowTextBox.Text, out var sampleRow, "Dòng mẫu"))
        {
            return;
        }

        _viewModel.ApplySingleSheetConfig(
            headerRow,
            nameColumn,
            sampleRow,
            AutoSkipBlankCheckBox.IsChecked == true,
            SuggestionsCheckBox.IsChecked == true,
            AutoSaveCheckBox.IsChecked == true);

        DialogResult = true;
        Close();
    }

    private void BackupNow_Click(object sender, RoutedEventArgs e)
    {
        _viewModel.ExportConfigBackup();
        MessageBox.Show(
            $"Đã lưu backup tại:\n{_viewModel.ConfigBackupFilePath}",
            "Backup cấu hình",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private static bool TryReadPositiveInt(string text, out int value, string fieldName)
    {
        if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && value >= 1)
        {
            return true;
        }

        MessageBox.Show($"{fieldName} phải là số >= 1.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        value = 0;
        return false;
    }

    private static bool TryReadNonNegativeInt(string text, out int value, string fieldName)
    {
        if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && value >= 0)
        {
            return true;
        }

        MessageBox.Show($"{fieldName} phải là số >= 0.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
        value = 0;
        return false;
    }
}
