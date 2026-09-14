using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class SortRecordsWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ObservableCollection<SortCustomValueItem> _customValues = [];

    public SortRecordsWindow(MainViewModel viewModel)
    {
        _viewModel = viewModel;
        InitializeComponent();
        SortColumnComboBox.ItemsSource = _viewModel.HeaderColumns;
        SortColumnComboBox.SelectedItem = _viewModel.HeaderColumns.FirstOrDefault();
        CustomValuesListBox.ItemsSource = _customValues;
        ReloadCustomValues();
        UpdateCustomPanelVisibility();
    }

    private void SortColumnComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ReloadCustomValues();
    }

    private void SortModeRadio_OnChecked(object sender, RoutedEventArgs e)
    {
        UpdateCustomPanelVisibility();
    }

    private void UpdateCustomPanelVisibility()
    {
        if (CustomOrderPanel is null || CustomRadio is null)
        {
            return;
        }

        CustomOrderPanel.Visibility = CustomRadio.IsChecked == true
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private void ReloadCustomValues()
    {
        _customValues.Clear();
        if (SortColumnComboBox.SelectedItem is not HeaderDefinition header)
        {
            return;
        }

        foreach (var value in _viewModel.GetDistinctSortColumnValues(header))
        {
            _customValues.Add(new SortCustomValueItem { Value = value });
        }
    }

    private void MoveUp_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: SortCustomValueItem item })
        {
            return;
        }

        var index = _customValues.IndexOf(item);
        if (index <= 0)
        {
            return;
        }

        _customValues.Move(index, index - 1);
    }

    private void MoveDown_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: SortCustomValueItem item })
        {
            return;
        }

        var index = _customValues.IndexOf(item);
        if (index < 0 || index >= _customValues.Count - 1)
        {
            return;
        }

        _customValues.Move(index, index + 1);
    }

    private void Sort_Click(object sender, RoutedEventArgs e)
    {
        if (SortColumnComboBox.SelectedItem is not HeaderDefinition header)
        {
            MessageBox.Show("Vui lòng chọn cột cần sắp xếp.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var mode = AscendingRadio.IsChecked == true
            ? ColumnSortMode.Ascending
            : DescendingRadio.IsChecked == true
                ? ColumnSortMode.Descending
                : ColumnSortMode.Custom;

        IReadOnlyList<string>? customOrder = null;
        if (mode == ColumnSortMode.Custom)
        {
            customOrder = _customValues.Select(v => v.Value).ToList();
            if (customOrder.Count == 0)
            {
                MessageBox.Show("Không có giá trị nào trong cột để sắp xếp tùy chỉnh.", "Thiếu thông tin", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        var error = _viewModel.SortRecordsByColumn(header, mode, customOrder);
        if (!string.IsNullOrWhiteSpace(error))
        {
            MessageBox.Show(error, "Sắp xếp thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        DialogResult = true;
        Close();
    }
}
