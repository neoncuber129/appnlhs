using System.Windows;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class MultiSheetConfigWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MultiSheetConfigWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        BackupPathTextBlock.Text = $"Backup tự động: {_viewModel.ConfigBackupFilePath}";
    }

    private void OpenImportConfig_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_viewModel.FilePath))
        {
            MessageBox.Show("Hãy mở file Excel trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new MultiSheetImportWindow(_viewModel)
        {
            Owner = this,
            Title = "Cấu hình nhập nhiều sheet"
        };

        if (dialog.ShowDialog() == true && dialog.ResultSession is not null)
        {
            _viewModel.ApplyMultiSheetImport(dialog.ResultSession);
        }
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

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
