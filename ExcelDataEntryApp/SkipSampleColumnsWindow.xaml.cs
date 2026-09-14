using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class SkipSampleColumnsWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly Func<MainViewModel, IReadOnlyList<HsskFemaleColumnOption>> _loadOptions;
    private readonly Func<MainViewModel, IEnumerable<int>, bool> _saveOptions;
    private readonly Func<MainViewModel, bool> _persistToAppFolder;
    private readonly Func<MainViewModel, string> _getProfilesFilePath;
    private readonly ObservableCollection<HsskFemaleColumnOption> _columns = [];
    private bool _isLoading;
    private static readonly Brush SavedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#6B7280")!);
    private static readonly Brush UnsavedBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B45309")!);

    public SkipSampleColumnsWindow(
        MainViewModel viewModel,
        string title,
        string helpText,
        Func<MainViewModel, IReadOnlyList<HsskFemaleColumnOption>> loadOptions,
        Func<MainViewModel, IEnumerable<int>, bool> saveOptions,
        Func<MainViewModel, bool> persistToAppFolder,
        Func<MainViewModel, string> getProfilesFilePath)
    {
        _viewModel = viewModel;
        _loadOptions = loadOptions;
        _saveOptions = saveOptions;
        _persistToAppFolder = persistToAppFolder;
        _getProfilesFilePath = getProfilesFilePath;
        InitializeComponent();
        Title = title;
        HelpTextBlock.Text = helpText;
        ColumnsItemsControl.ItemsSource = _columns;
        ReloadColumns();
        UpdateSavedHint(saved: false);
    }

    private void ReloadColumns()
    {
        _isLoading = true;
        try
        {
            _columns.Clear();
            foreach (var item in _loadOptions(_viewModel))
            {
                item.PropertyChanged += ColumnOption_PropertyChanged;
                _columns.Add(item);
            }
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void ColumnOption_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_isLoading || e.PropertyName != nameof(HsskFemaleColumnOption.IsSelected))
        {
            return;
        }

        PersistProfile();
    }

    private void PersistProfile(bool showSavedMessage = false)
    {
        var saved = _saveOptions(_viewModel, _columns.Where(c => c.IsSelected).Select(c => c.ColumnIndex));
        if (saved)
        {
            saved = _persistToAppFolder(_viewModel);
        }

        UpdateSavedHint(saved);
        if (showSavedMessage && saved)
        {
            MessageBox.Show(
                $"Đã lưu profile cấu hình cột $.\n\n"
                + $"File profile:\n{_getProfilesFilePath(_viewModel)}\n\n"
                + $"File backup:\n{_viewModel.AppConfigBackupFilePath}",
                "Hoàn tất",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        if (!_viewModel.HasExportProfileKey)
        {
            MessageBox.Show(
                "Chưa lưu được — cần mở file Excel và áp dụng tiêu đề trước.",
                "Thiếu thông tin",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            UpdateSavedHint(saved: false);
            return;
        }

        PersistProfile(showSavedMessage: true);
    }

    private void UpdateSavedHint(bool saved)
    {
        if (saved)
        {
            SavedHintTextBlock.Text =
                $"Đã lưu ({DateTime.Now:HH:mm:ss}) — {_getProfilesFilePath(_viewModel)}";
            SavedHintTextBlock.Foreground = SavedBrush;
            return;
        }

        SavedHintTextBlock.Text = _viewModel.HasExportProfileKey
            ? "Chưa lưu — bấm Lưu profile"
            : "Chưa lưu — cần mở file Excel và áp dụng tiêu đề";
        SavedHintTextBlock.Foreground = UnsavedBrush;
    }

    private void ScanMarkerColumns_Click(object sender, RoutedEventArgs e)
    {
        var scanned = _viewModel.ScanMarkerColumnIndexes().ToHashSet();
        if (scanned.Count == 0)
        {
            MessageBox.Show(
                "Không tìm thấy cột nào có ký hiệu \"s\" ở dòng 1 hoặc dòng 2.\n\n"
                + "Bạn vẫn có thể tick chọn thủ công từng cột bên dưới.",
                "Gợi ý cột",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        foreach (var item in _columns.Where(c => scanned.Contains(c.ColumnIndex)))
        {
            item.IsSelected = true;
        }

        PersistProfile();
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _columns)
        {
            item.IsSelected = true;
        }

        PersistProfile();
    }

    private void UnselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var item in _columns)
        {
            item.IsSelected = false;
        }

        PersistProfile();
    }
}
