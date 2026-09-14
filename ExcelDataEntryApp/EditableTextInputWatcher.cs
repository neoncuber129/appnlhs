using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp;

/// <summary>
/// TextBox nhập liệu: cập nhật model khi gõ nhưng không làm binding đẩy ngược Text (giữ vị trí con trỏ).
/// </summary>
public static class EditableTextInputWatcher
{
    public static readonly DependencyProperty EnableProperty = DependencyProperty.RegisterAttached(
        "Enable",
        typeof(bool),
        typeof(EditableTextInputWatcher),
        new PropertyMetadata(false, OnEnableChanged));

    public static void SetEnable(TextBox element, bool value) => element.SetValue(EnableProperty, value);

    public static bool GetEnable(TextBox element) => (bool)element.GetValue(EnableProperty);

    private sealed class Hook
    {
        public CaretRestoreGate CaretGate { get; } = new();
        public TextChangedEventHandler TextChanged { get; set; } = null!;
        public KeyboardFocusChangedEventHandler LostKeyboardFocus { get; set; } = null!;
    }

    private static void OnEnableChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not TextBox textBox)
        {
            return;
        }

        if ((bool)e.NewValue)
        {
            textBox.Loaded += TextBox_Loaded;
            if (textBox.IsLoaded)
            {
                TextBox_Loaded(textBox, new RoutedEventArgs());
            }
        }
        else
        {
            textBox.Loaded -= TextBox_Loaded;
            Unhook(textBox);
        }
    }

    private static void TextBox_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not TextBox textBox || textBox.Tag is Hook)
        {
            return;
        }

        var hook = new Hook();
        hook.TextChanged = (_, _) => OnTextChanged(textBox, hook);
        hook.LostKeyboardFocus = (_, _) => OnLostKeyboardFocus(textBox);
        textBox.Tag = hook;
        TextEditorCaretHelper.AttachCaretRestoreGate(textBox, hook.CaretGate);
        textBox.TextChanged += hook.TextChanged;
        textBox.LostKeyboardFocus += hook.LostKeyboardFocus;
        textBox.Unloaded += TextBox_Unloaded;
    }

    private static void OnTextChanged(TextBox textBox, Hook hook)
    {
        if (textBox.DataContext is not EditableField field)
        {
            return;
        }

        var text = textBox.Text ?? string.Empty;
        var caret = textBox.CaretIndex;
        var selectionLength = textBox.SelectionLength;
        var token = hook.CaretGate.Next();
        field.SetValueSilently(text);
        TextEditorCaretHelper.RestoreTextBoxCaret(textBox, caret, selectionLength, hook.CaretGate, token);
    }

    private static void OnLostKeyboardFocus(TextBox textBox)
    {
        if (textBox.DataContext is not EditableField field)
        {
            return;
        }

        field.CommitValueFromEditor(textBox.Text ?? string.Empty);
    }

    private static void TextBox_Unloaded(object sender, RoutedEventArgs e) => Unhook(sender as TextBox);

    private static void Unhook(TextBox? textBox)
    {
        if (textBox?.Tag is not Hook hook)
        {
            return;
        }

        textBox.TextChanged -= hook.TextChanged;
        textBox.LostKeyboardFocus -= hook.LostKeyboardFocus;
        textBox.Unloaded -= TextBox_Unloaded;
        textBox.Tag = null;
    }
}
