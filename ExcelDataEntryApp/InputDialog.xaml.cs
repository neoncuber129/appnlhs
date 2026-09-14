using System.Windows;

namespace ExcelDataEntryApp;

public partial class InputDialog : Window
{
    public string InputValue { get; private set; } = string.Empty;

    public InputDialog(string prompt = "Nhập liệu:", string defaultValue = "", string title = "Nhập liệu")
    {
        InitializeComponent();
        Title = title;
        PromptText.Text = prompt;
        InputTextBox.Text = defaultValue;
        InputTextBox.Focus();
        InputTextBox.SelectAll();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        InputValue = InputTextBox.Text;
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        InputValue = string.Empty;
        DialogResult = false;
        Close();
    }
}
