using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel();
        DataContext = _viewModel;
        _viewModel.PropertyChanged += ViewModelOnPropertyChanged;
        Loaded += (_, _) => SyncMultiSheetModeToggle();
        AddHandler(ComboBox.PreviewMouseWheelEvent, new MouseWheelEventHandler(ComboBox_PreviewMouseWheel), true);
        AddHandler(Keyboard.PreviewKeyDownEvent, new KeyEventHandler(InputField_PreviewKeyDown), true);
        PreviewMouseDown += MainWindow_PreviewMouseDown;
    }

    private void MainWindow_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.OriginalSource is DependencyObject src)
        {
            SuggestionComboBoxWatcher.CloseDropDownsIfClickOutside(this, src);
            DropdownComboBoxWatcher.CloseDropDownsIfClickOutside(this, src);
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        _viewModel.PropertyChanged -= ViewModelOnPropertyChanged;
        _viewModel.Dispose();
        base.OnClosed(e);
    }

    private void ViewModelOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.IsMultiSheetMode))
        {
            SyncMultiSheetModeToggle();
            return;
        }

        if (e.PropertyName != nameof(MainViewModel.SelectedRecord))
        {
            return;
        }

        Dispatcher.BeginInvoke(new Action(ScrollFieldsPanelToTop), DispatcherPriority.Loaded);
    }

    private void ScrollFieldsPanelToTop()
    {
        FieldsScrollViewer.ScrollToHome();
        MultiFieldsScrollViewer.ScrollToHome();
    }

    private void InputField_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || sender is not Control control)
        {
            return;
        }

        e.Handled = true;
        control.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
    }

    private void InputField_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter && e.Key != Key.Down && e.Key != Key.Up)
        {
            return;
        }

        if (e.OriginalSource is not DependencyObject source)
        {
            return;
        }

        var control = FindParentControl(source);
        if (control is null || control.DataContext is not EditableField)
        {
            return;
        }

        if (control is ComboBox comboBox && comboBox.IsDropDownOpen && (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Enter))
        {
            return;
        }

        var direction = e.Key == Key.Up
            ? FocusNavigationDirection.Previous
            : FocusNavigationDirection.Next;
        e.Handled = true;
        control.MoveFocus(new TraversalRequest(direction));
    }

    private void RecordsList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBox listBox)
        {
            return;
        }

        var hit = e.OriginalSource as DependencyObject;
        if (hit is null)
        {
            return;
        }

        if (listBox.ContainerFromElement(hit) is not ListBoxItem item || item.DataContext is not RecordItem record)
        {
            return;
        }

        // Một lần click: chọn đúng dòng, focus list, gán SelectedItem để binding cập nhật form ngay (tránh lệch hit-test / focus).
        if (!ReferenceEquals(listBox.SelectedItem, record))
        {
            listBox.SelectedItem = record;
        }

        if (!item.IsFocused)
        {
            item.Focus();
        }

        listBox.Focus();
    }

    private void RecordsList_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBox listBox)
        {
            return;
        }

        var hit = e.OriginalSource as DependencyObject;
        if (hit is null)
        {
            return;
        }

        if (listBox.ContainerFromElement(hit) is not ListBoxItem item || item.DataContext is not RecordItem record)
        {
            return;
        }

        if (!ReferenceEquals(listBox.SelectedItem, record))
        {
            listBox.SelectedItem = record;
        }

        if (!item.IsFocused)
        {
            item.Focus();
        }

        listBox.Focus();
    }

    private void OpenSingleSheetConfig_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SingleSheetConfigWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenMultiSheetConfig_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new MultiSheetConfigWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenHeaderSettings_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new HeaderSettingsWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenExportHssk_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ExportHsskWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenExportImportFile_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_viewModel.FilePath))
        {
            MessageBox.Show("Hãy mở file Excel trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (_viewModel.HeaderColumns.Count == 0)
        {
            MessageBox.Show("Chưa có tiêu đề cột. Hãy nhận diện header trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new ExportImportFileWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenSortRecords_Click(object sender, RoutedEventArgs e)
    {
        if (_viewModel.IsMultiSheetMode)
        {
            MessageBox.Show("Sắp xếp chưa hỗ trợ khi đang ghép nhiều sheet. Hãy tắt \"Nhập nhiều sheet\" trước.", "Chưa hỗ trợ", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (_viewModel.HeaderColumns.Count == 0)
        {
            MessageBox.Show("Chưa có tiêu đề cột. Hãy mở file Excel và nhận diện header trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new SortRecordsWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void MultiSheetModeToggle_Checked(object sender, RoutedEventArgs e)
    {
        if (_viewModel.IsMultiSheetMode)
        {
            return;
        }

        if (!TryEnableMultiSheetMode())
        {
            SyncMultiSheetModeToggle();
        }
    }

    private void MultiSheetModeToggle_Unchecked(object sender, RoutedEventArgs e)
    {
        if (!_viewModel.IsMultiSheetMode)
        {
            return;
        }

        _viewModel.ExitMultiSheetMode();
    }

    private bool TryEnableMultiSheetMode()
    {
        if (string.IsNullOrWhiteSpace(_viewModel.FilePath))
        {
            MessageBox.Show("Hãy mở file Excel trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        var dialog = new MultiSheetImportWindow(_viewModel)
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true && dialog.ResultSession is not null)
        {
            _viewModel.ApplyMultiSheetImport(dialog.ResultSession);
            return true;
        }

        return false;
    }

    private void SyncMultiSheetModeToggle()
    {
        if (MultiSheetModeToggle is null)
        {
            return;
        }

        MultiSheetModeToggle.IsChecked = _viewModel.IsMultiSheetMode;
    }

    private void OpenLinkData_Click(object sender, RoutedEventArgs e)
    {
        if (_viewModel.HeaderColumns.Count == 0)
        {
            MessageBox.Show("Chưa có tiêu đề cột. Hãy mở file Excel và nhận diện header trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new LinkDataWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void OpenImportListLink_Click(object sender, RoutedEventArgs e)
    {
        if (_viewModel.HeaderColumns.Count == 0)
        {
            MessageBox.Show("Chưa có tiêu đề cột. Hãy mở file Excel và nhận diện header trước.", "Chưa sẵn sàng", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new ImportListLinkWindow(_viewModel)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void ComboBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        var comboBox = FindParentComboBox(e.OriginalSource as DependencyObject);
        if (comboBox is not null && !comboBox.IsDropDownOpen)
        {
            e.Handled = true;
            var scrollViewer = FindParentScrollViewer(comboBox);
            if (scrollViewer is not null)
            {
                var args = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
                {
                    RoutedEvent = UIElement.MouseWheelEvent,
                    Source = comboBox
                };
                scrollViewer.RaiseEvent(args);
            }
        }
    }

    private static ComboBox? FindParentComboBox(DependencyObject? current)
    {
        while (current is not null)
        {
            if (current is ComboBox comboBox)
            {
                return comboBox;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static ScrollViewer? FindParentScrollViewer(DependencyObject? current)
    {
        while (current is not null)
        {
            if (current is ScrollViewer scrollViewer)
            {
                return scrollViewer;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private static Control? FindParentControl(DependencyObject? current)
    {
        while (current is not null)
        {
            if (current is Control control)
            {
                return control;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }
}