using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace ExcelDataEntryApp;

internal sealed class CaretRestoreGate
{
    private int _token;

    public int Next() => Interlocked.Increment(ref _token);

    public void CancelPending() => Interlocked.Increment(ref _token);

    public bool IsCurrent(int token) => token == Volatile.Read(ref _token);
}

internal static class TextEditorCaretHelper
{
    public static bool TryGetEditableTextBox(ComboBox comboBox, out TextBox? textBox)
    {
        textBox = null;
        if (!comboBox.IsEditable)
        {
            return false;
        }

        comboBox.ApplyTemplate();
        textBox = comboBox.Template?.FindName("PART_EditableTextBox", comboBox) as TextBox;
        return textBox is not null;
    }

    public static void AttachCaretRestoreGate(UIElement element, CaretRestoreGate gate)
    {
        element.AddHandler(
            UIElement.PreviewMouseDownEvent,
            new MouseButtonEventHandler((_, _) => gate.CancelPending()),
            handledEventsToo: true);
    }

    public static void PreserveEditableComboCaret(ComboBox comboBox, CaretRestoreGate gate, Action update)
    {
        if (!TryGetEditableTextBox(comboBox, out var textBox) || textBox is null)
        {
            update();
            return;
        }

        var caret = textBox.CaretIndex;
        var selectionLength = textBox.SelectionLength;
        var token = gate.Next();
        update();
        RestoreTextBoxCaret(textBox, caret, selectionLength, gate, token);
    }

    public static void RestoreTextBoxCaret(
        TextBox textBox,
        int caretIndex,
        int selectionLength,
        CaretRestoreGate gate,
        int token)
    {
        textBox.Dispatcher.BeginInvoke(
            () =>
            {
                if (!gate.IsCurrent(token) || !textBox.IsKeyboardFocusWithin)
                {
                    return;
                }

                var length = textBox.Text?.Length ?? 0;
                var caret = Math.Clamp(caretIndex, 0, length);
                textBox.CaretIndex = caret;
                textBox.SelectionLength = Math.Clamp(selectionLength, 0, length - caret);
            },
            DispatcherPriority.Input);
    }
}
