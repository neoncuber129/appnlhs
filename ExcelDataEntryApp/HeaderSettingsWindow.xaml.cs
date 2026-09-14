using System.Windows;
using ExcelDataEntryApp.Models;
using ExcelDataEntryApp.ViewModels;

namespace ExcelDataEntryApp;

public partial class HeaderSettingsWindow : Window
{
    public HeaderSettingsWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ToggleHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: HeaderDefinition header })
        {
            header.IsVisible = !header.IsVisible;
        }
    }
}
