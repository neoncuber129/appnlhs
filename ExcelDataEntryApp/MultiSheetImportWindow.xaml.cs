using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.Services;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class MultiSheetImportWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ObservableCollection<SheetImportConfigItem> _items = [];

    public MultiSheetImportSession? ResultSession { get; private set; }

    public MultiSheetImportWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = _viewModel;
        var saved = _viewModel.TryGetSavedMultiSheetConfig();
        var includeHidden = saved?.ShowHiddenSheets ?? false;
        if (includeHidden)
        {
            ShowHiddenSheetsCheckBox.IsChecked = true;
        }

        AutoSkipBlankCheckBox.IsChecked = saved?.AutoSkipBlankHeaders ?? MultiSheetImportDefaults.AutoSkipBlankHeaders;
        SuggestionsCheckBox.IsChecked = saved is null
            ? !MultiSheetImportDefaults.SuggestionsDisabled
            : !saved.SuggestionsDisabled;

        LoadSheets(includeHidden, saved);
    }

    private void LoadSheets(bool includeHidden, SavedMultiSheetConfig? saved = null)
    {
        _items.Clear();
        var sheetDefaults = ResolveSheetDefaults(saved);
        var savedBySheet = saved?.Sheets.ToDictionary(s => s.SheetName, StringComparer.Ordinal) ?? [];

        foreach (var info in _viewModel.GetWorksheetInfos(includeHidden))
        {
            savedBySheet.TryGetValue(info.Name, out var savedSheet);
            var item = new SheetImportConfigItem
            {
                SheetName = info.Name,
                IsHidden = info.IsHidden,
                IsSelected = savedSheet is not null || !info.IsHidden,
                HeaderFirstRow = savedSheet?.HeaderFirstRow ?? sheetDefaults.HeaderFirstRow,
                HeaderLastRow = savedSheet?.HeaderLastRow ?? sheetDefaults.HeaderLastRow,
                FirstDataRow = savedSheet?.FirstDataRow ?? sheetDefaults.FirstDataRow,
                SampleRow = savedSheet?.SampleRow > 0
                    ? savedSheet.SampleRow
                    : (savedSheet?.FirstDataRow > 0 ? savedSheet.FirstDataRow : sheetDefaults.SampleRow),
                NameColumnIndex = savedSheet?.NameColumnIndex ?? sheetDefaults.NameColumnIndex,
                IsNameSheet = savedSheet?.IsNameSheet ?? false
            };
            item.RefreshColumnOptionsRequested += (_, _) => RefreshColumnOptions(item);
            RefreshColumnOptions(item);
            _items.Add(item);
        }

        if (saved is not null)
        {
            foreach (var item in _items)
            {
                item.IsSelected = saved.Sheets.Any(s => string.Equals(s.SheetName, item.SheetName, StringComparison.Ordinal));
            }

            var nameItem = _items.FirstOrDefault(i => string.Equals(i.SheetName, saved.NameSheetName, StringComparison.Ordinal));
            if (nameItem is not null)
            {
                foreach (var item in _items)
                {
                    item.IsNameSheet = ReferenceEquals(item, nameItem);
                }

                nameItem.NameColumnIndex = saved.NameColumnIndex;
            }
        }
        else if (_items.Count > 0 && !_items.Any(i => i.IsNameSheet))
        {
            _items[0].IsNameSheet = true;
        }

        SheetsDataGrid.ItemsSource = _items;
    }

    private void RefreshColumnOptions(SheetImportConfigItem item)
    {
        item.ColumnOptions.Clear();
        foreach (var header in _viewModel.ReadHeaderBlock(item.SheetName, item.HeaderFirstRow, item.HeaderLastRow))
        {
            item.ColumnOptions.Add(new HeaderColumnOption
            {
                ColumnIndex = header.ColumnIndex,
                DisplayName = header.Name
            });
        }

        if (item.ColumnOptions.All(c => c.ColumnIndex != item.NameColumnIndex))
        {
            item.NameColumnIndex = item.ColumnOptions.FirstOrDefault()?.ColumnIndex ?? _viewModel.NameColumnIndex;
        }
    }

    private void RefreshAllColumnOptions()
    {
        foreach (var item in _items)
        {
            RefreshColumnOptions(item);
        }
    }

    private void ShowHiddenSheets_Changed(object sender, RoutedEventArgs e)
    {
        var includeHidden = ShowHiddenSheetsCheckBox.IsChecked == true;
        var previous = _items.ToDictionary(i => i.SheetName, i => i);
        LoadSheets(includeHidden);

        foreach (var item in _items)
        {
            if (!previous.TryGetValue(item.SheetName, out var old))
            {
                continue;
            }

            item.IsSelected = old.IsSelected;
            item.HeaderFirstRow = old.HeaderFirstRow;
            item.HeaderLastRow = old.HeaderLastRow;
            item.FirstDataRow = old.FirstDataRow;
            item.SampleRow = old.SampleRow;
            item.IsNameSheet = old.IsNameSheet;
            item.NameColumnIndex = old.NameColumnIndex;
            RefreshColumnOptions(item);
        }
    }

    private void SheetsDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.Row.Item is SheetImportConfigItem item)
        {
            RefreshColumnOptions(item);
        }
    }

    private void NameSheetRadio_Checked(object sender, RoutedEventArgs e)
    {
        if (sender is not RadioButton radio || radio.DataContext is not SheetImportConfigItem selected)
        {
            return;
        }

        foreach (var item in _items)
        {
            item.IsNameSheet = ReferenceEquals(item, selected);
            if (item.IsNameSheet)
            {
                item.IsSelected = true;
            }
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        SheetsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        SheetsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        RefreshAllColumnOptions();

        var selected = _items.Where(i => i.IsSelected).ToList();
        if (selected.Count < 2)
        {
            MessageBox.Show("Vui lòng chọn ít nhất 2 sheet.", "Thiếu sheet", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var nameSheets = selected.Where(i => i.IsNameSheet).ToList();
        if (nameSheets.Count != 1)
        {
            MessageBox.Show("Vui lòng chọn đúng 1 sheet chứa cột tên.", "Sheet tên", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var nameSheet = nameSheets[0];
        if (nameSheet.ColumnOptions.All(c => c.ColumnIndex != nameSheet.NameColumnIndex))
        {
            MessageBox.Show("Cột tên không hợp lệ trên sheet đã chọn.", "Cột tên", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        foreach (var item in selected)
        {
            if (!int.TryParse(item.HeaderFirstRow.ToString(CultureInfo.InvariantCulture), out var headerFirst) || headerFirst < 1)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': tiêu đề từ hàng phải là số >= 1.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(item.HeaderLastRow.ToString(CultureInfo.InvariantCulture), out var headerLast) || headerLast < headerFirst)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': tiêu đề đến hàng phải >= tiêu đề từ hàng ({headerFirst}).", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(item.FirstDataRow.ToString(CultureInfo.InvariantCulture), out var firstData) || firstData <= headerLast)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': dòng bắt đầu nhập phải lớn hơn tiêu đề đến hàng ({headerLast}).", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(item.SampleRow.ToString(CultureInfo.InvariantCulture), out var sampleRow) || sampleRow < 1)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': dòng mẫu phải là số >= 1.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (sampleRow <= headerLast)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': dòng mẫu phải lớn hơn tiêu đề đến hàng ({headerLast}).", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (sampleRow > firstData)
            {
                MessageBox.Show($"Sheet '{item.SheetName}': dòng mẫu không được lớn hơn dòng bắt đầu nhập ({firstData}).", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        ResultSession = new MultiSheetImportSession
        {
            Sheets = selected.Select(i => i.ToConfig()).ToList(),
            NameSheetName = nameSheet.SheetName,
            NameColumnIndex = nameSheet.NameColumnIndex,
            ShowHiddenSheets = ShowHiddenSheetsCheckBox.IsChecked == true,
            AutoSkipBlankHeaders = AutoSkipBlankCheckBox.IsChecked == true,
            SuggestionsDisabled = SuggestionsCheckBox.IsChecked != true
        };

        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private static (int HeaderFirstRow, int HeaderLastRow, int FirstDataRow, int SampleRow, int NameColumnIndex) ResolveSheetDefaults(
        SavedMultiSheetConfig? saved)
    {
        if (saved?.Sheets.Count > 0)
        {
            var reference = saved.Sheets[0];
            return (
                reference.HeaderFirstRow,
                reference.HeaderLastRow,
                reference.FirstDataRow,
                reference.SampleRow > 0 ? reference.SampleRow : reference.FirstDataRow,
                saved.NameColumnIndex);
        }

        return (
            MultiSheetImportDefaults.HeaderFirstRow,
            MultiSheetImportDefaults.HeaderLastRow,
            MultiSheetImportDefaults.FirstDataRow,
            MultiSheetImportDefaults.SampleRow,
            MultiSheetImportDefaults.NameColumnIndex);
    }
}
