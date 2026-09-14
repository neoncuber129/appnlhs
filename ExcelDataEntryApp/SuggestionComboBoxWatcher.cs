using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp;

/// <summary>
/// ComboBox gợi ý: bảng gợi ý chỉ mở một lần khi vào ô (focus); kể từ khi text khác lúc vào ô thì không mở lại (chỉ dùng chữ người dùng gõ).
/// Sau khi xóa hết từ lúc đang có chữ thì lần sau vào ô trống không tự bung list (SuppressDropDownWhileEmpty).
/// Xóa SelectedItem khi text rỗng hoặc khi text khác mục đang chọn — tránh WPF đồng bộ Text ngược về gợi ý cũ khi gõ sửa.
/// </summary>
public static class SuggestionComboBoxWatcher
{
    public static readonly DependencyProperty EnableProperty = DependencyProperty.RegisterAttached(
        "Enable",
        typeof(bool),
        typeof(SuggestionComboBoxWatcher),
        new PropertyMetadata(false, OnEnableChanged));

    public static void SetEnable(ComboBox element, bool value) => element.SetValue(EnableProperty, value);

    public static bool GetEnable(ComboBox element) => (bool)element.GetValue(EnableProperty);

    /// <summary>
    /// Đóng dropdown gợi ý khi click không thuộc ComboBox / popup của nó — để một cú click chuyển được sang ô khác.
    /// Gọi từ PreviewMouseDown trên Window (tunnel, đầu tiên).
    /// </summary>
    public static void CloseDropDownsIfClickOutside(Window window, DependencyObject? originalSource)
    {
        if (originalSource is null)
        {
            return;
        }

        var list = new List<ComboBox>();
        CollectComboBoxes(window, list);

        foreach (var cb in list)
        {
            if (!GetEnable(cb) || !cb.IsDropDownOpen)
            {
                continue;
            }

            if (IsClickInsideComboBoxScope(cb, originalSource))
            {
                continue;
            }

            // Popup list đôi khi không nối visual lên ComboBox — vẫn coi là click trong scope.
            if (IsOriginalSourceUnderComboPopup(cb, originalSource))
            {
                continue;
            }

            cb.IsDropDownOpen = false;
        }
    }

    private static bool IsOriginalSourceUnderComboPopup(ComboBox cb, DependencyObject src)
    {
        if (cb.Template?.FindName("PART_Popup", cb) is not Popup { Child: Visual popupRoot })
        {
            return false;
        }

        for (var cur = src; cur is not null; cur = VisualTreeHelper.GetParent(cur))
        {
            if (ReferenceEquals(cur, popupRoot))
            {
                return true;
            }
        }

        return false;
    }

    private static void CollectComboBoxes(DependencyObject node, List<ComboBox> list)
    {
        if (node is ComboBox cb)
        {
            list.Add(cb);
        }

        var n = VisualTreeHelper.GetChildrenCount(node);
        for (var i = 0; i < n; i++)
        {
            CollectComboBoxes(VisualTreeHelper.GetChild(node, i), list);
        }
    }

    private static bool IsClickInsideComboBoxScope(ComboBox cb, DependencyObject src)
    {
        for (var cur = src; cur is not null; cur = VisualTreeHelper.GetParent(cur))
        {
            if (ReferenceEquals(cur, cb))
            {
                return true;
            }

            if (cur is FrameworkElement fe && ReferenceEquals(fe.TemplatedParent, cb))
            {
                return true;
            }

            if (cur is Popup { PlacementTarget: DependencyObject pt })
            {
                if (ReferenceEquals(pt, cb))
                {
                    return true;
                }

                for (var p = pt; p is not null; p = VisualTreeHelper.GetParent(p))
                {
                    if (ReferenceEquals(p, cb))
                    {
                        return true;
                    }

                    if (p is FrameworkElement pfe && ReferenceEquals(pfe.TemplatedParent, cb))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private sealed class Hook
    {
        public CaretRestoreGate CaretGate { get; } = new();
        public TextChangedEventHandler TextChanged { get; set; } = null!;
        public SelectionChangedEventHandler SelectionChanged { get; set; } = null!;
        public EventHandler DropDownClosed { get; set; } = null!;
        public KeyboardFocusChangedEventHandler GotKeyboardFocus { get; set; } = null!;
        public KeyboardFocusChangedEventHandler LostKeyboardFocus { get; set; } = null!;
        public string LastKnownText { get; set; } = "";
        public string TextAtFocusSnapshot { get; set; } = "";
        public bool UserHasChangedTextSinceFocus { get; set; }
        public bool SuppressDropDownWhileEmpty { get; set; }
    }

    private static void OnEnableChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not ComboBox cb)
        {
            return;
        }

        if ((bool)e.NewValue)
        {
            cb.Loaded += ComboBoxHook_Loaded;
            if (cb.IsLoaded)
            {
                ComboBoxHook_Loaded(cb, new RoutedEventArgs());
            }
        }
        else
        {
            cb.Loaded -= ComboBoxHook_Loaded;
            Unhook(cb);
        }
    }

    private static void ComboBoxHook_Loaded(object sender, RoutedEventArgs e)
    {
        if (sender is not ComboBox cb || cb.Tag is Hook)
        {
            return;
        }

        var hook = new Hook
        {
            LastKnownText = cb.Text ?? string.Empty,
            SuppressDropDownWhileEmpty = false
        };
        hook.TextChanged = (_, _) => OnSuggestionTextChanged(cb, hook);
        hook.SelectionChanged = (_, e) => OnSuggestionSelectionChanged(cb, hook, e);
        hook.DropDownClosed = (_, _) => OnSuggestionDropDownClosed(cb, hook);
        hook.GotKeyboardFocus = (_, _) =>
        {
            hook.UserHasChangedTextSinceFocus = false;
            hook.TextAtFocusSnapshot = cb.Text ?? string.Empty;
            hook.LastKnownText = hook.TextAtFocusSnapshot;
            SyncDropDown(cb, hook);
        };
        hook.LostKeyboardFocus = (_, _) =>
        {
            cb.IsDropDownOpen = false;
            if (cb.DataContext is EditableField field)
            {
                field.CommitValueFromEditor(cb.Text ?? string.Empty);
            }
        };

        cb.Tag = hook;
        TextEditorCaretHelper.AttachCaretRestoreGate(cb, hook.CaretGate);
        cb.AddHandler(TextBoxBase.TextChangedEvent, hook.TextChanged);
        cb.SelectionChanged += hook.SelectionChanged;
        cb.DropDownClosed += hook.DropDownClosed;
        cb.GotKeyboardFocus += hook.GotKeyboardFocus;
        cb.LostKeyboardFocus += hook.LostKeyboardFocus;
        cb.Unloaded += ComboBoxHook_Unloaded;
    }

    private static void OnSuggestionSelectionChanged(ComboBox cb, Hook hook, SelectionChangedEventArgs e)
    {
        if (!cb.IsEditable || e.AddedItems.Count == 0)
        {
            return;
        }

        var selected = e.AddedItems[0] as string ?? Convert.ToString(e.AddedItems[0]) ?? string.Empty;
        if (string.IsNullOrEmpty(selected))
        {
            return;
        }

        CommitSuggestionSelection(cb, hook, selected);
    }

    private static void OnSuggestionDropDownClosed(ComboBox cb, Hook hook)
    {
        if (cb.IsEditable
            && cb.SelectedItem is not null
            && !string.IsNullOrWhiteSpace(cb.Text))
        {
            var selected = cb.SelectedItem as string ?? Convert.ToString(cb.SelectedItem) ?? string.Empty;
            if (!string.IsNullOrEmpty(selected))
            {
                CommitSuggestionSelection(cb, hook, selected);
            }
        }

        ClearSelectionIfTextEmpty(cb);
    }

    private static void CommitSuggestionSelection(ComboBox cb, Hook hook, string selected)
    {
        if (hook.UserHasChangedTextSinceFocus
            && string.Equals(hook.LastKnownText, selected, StringComparison.Ordinal)
            && string.Equals(hook.TextAtFocusSnapshot, selected, StringComparison.Ordinal))
        {
            cb.Dispatcher.BeginInvoke(() => PlaceCaretAtEnd(cb), DispatcherPriority.Input);
            return;
        }

        hook.UserHasChangedTextSinceFocus = true;
        hook.LastKnownText = selected;
        hook.TextAtFocusSnapshot = selected;
        hook.CaretGate.CancelPending();

        if (cb.DataContext is EditableField field)
        {
            field.SetValueSilently(selected);
        }

        cb.Dispatcher.BeginInvoke(
            () =>
            {
                if (!cb.IsKeyboardFocusWithin)
                {
                    return;
                }

                cb.SelectedItem = null;
                if (!string.Equals(cb.Text, selected, StringComparison.Ordinal))
                {
                    cb.Text = selected;
                }

                PlaceCaretAtEnd(cb);
            },
            DispatcherPriority.Input);
    }

    private static void PlaceCaretAtEnd(ComboBox cb)
    {
        if (!TextEditorCaretHelper.TryGetEditableTextBox(cb, out var editableTextBox) || editableTextBox is null)
        {
            return;
        }

        var length = editableTextBox.Text?.Length ?? 0;
        editableTextBox.CaretIndex = length;
        editableTextBox.SelectionLength = 0;
    }

    private static void OnSuggestionTextChanged(ComboBox cb, Hook hook)
    {
        var cur = cb.Text ?? string.Empty;
        var prev = hook.LastKnownText;
        _ = TextEditorCaretHelper.TryGetEditableTextBox(cb, out var editableTextBox);
        var caretBeforeDetach = editableTextBox?.CaretIndex ?? cur.Length;

        var hadSelectedItem = cb.SelectedItem is not null;
        DetachSelectedItemIfTextMismatch(cb, cur);

        if (!string.Equals(cur, hook.TextAtFocusSnapshot, StringComparison.Ordinal))
        {
            hook.UserHasChangedTextSinceFocus = true;
        }

        if (!string.IsNullOrWhiteSpace(prev) && string.IsNullOrWhiteSpace(cur))
        {
            hook.SuppressDropDownWhileEmpty = true;
        }

        if (!string.IsNullOrWhiteSpace(cur))
        {
            hook.SuppressDropDownWhileEmpty = false;
        }

        hook.LastKnownText = cur;

        if (cb.DataContext is EditableField field)
        {
            field.SetValueSilently(cur);
        }

        if (editableTextBox is not null && hadSelectedItem)
        {
            var targetCaret = cur.Length > prev.Length && caretBeforeDetach >= prev.Length
                ? cur.Length
                : caretBeforeDetach;
            if (editableTextBox.CaretIndex != targetCaret)
            {
                var token = hook.CaretGate.Next();
                TextEditorCaretHelper.RestoreTextBoxCaret(editableTextBox, targetCaret, 0, hook.CaretGate, token);
            }
        }

        ClearSelectionIfTextEmpty(cb);

        if (hook.UserHasChangedTextSinceFocus)
        {
            cb.IsDropDownOpen = false;
        }
    }

    private static void SyncDropDown(ComboBox cb, Hook hook)
    {
        if (!cb.IsKeyboardFocusWithin)
        {
            cb.IsDropDownOpen = false;
            return;
        }

        if (hook.UserHasChangedTextSinceFocus)
        {
            cb.IsDropDownOpen = false;
            return;
        }

        var empty = string.IsNullOrWhiteSpace(cb.Text);
        if (empty && hook.SuppressDropDownWhileEmpty)
        {
            cb.IsDropDownOpen = false;
            return;
        }

        cb.IsDropDownOpen = true;
    }

    /// <summary>Gỡ neo SelectedItem khi người dùng gõ khác chuỗi đã chọn trong list (editable ComboBox hay kéo Text về item cũ).</summary>
    private static void DetachSelectedItemIfTextMismatch(ComboBox cb, string cur)
    {
        if (!cb.IsEditable || cb.SelectedItem is null)
        {
            return;
        }

        var selectedText = cb.SelectedItem as string ?? Convert.ToString(cb.SelectedItem) ?? string.Empty;
        if (string.Equals(cur, selectedText, StringComparison.Ordinal))
        {
            return;
        }

        cb.SelectedItem = null;
        if (!string.Equals(cb.Text, cur, StringComparison.Ordinal))
        {
            cb.Text = cur;
        }
    }

    private static void ClearSelectionIfTextEmpty(ComboBox cb)
    {
        if (!cb.IsEditable || !string.IsNullOrWhiteSpace(cb.Text))
        {
            return;
        }

        cb.SelectedItem = null;
        if (cb.DataContext is EditableField f && f.Value.Length > 0)
        {
            f.SetValueSilently(string.Empty);
        }
    }

    private static void ComboBoxHook_Unloaded(object sender, RoutedEventArgs e)
    {
        Unhook(sender as ComboBox);
    }

    private static void Unhook(ComboBox? cb)
    {
        if (cb is null)
        {
            return;
        }

        if (cb.Tag is not Hook hook)
        {
            return;
        }

        cb.RemoveHandler(TextBoxBase.TextChangedEvent, hook.TextChanged);
        cb.SelectionChanged -= hook.SelectionChanged;
        cb.DropDownClosed -= hook.DropDownClosed;
        cb.GotKeyboardFocus -= hook.GotKeyboardFocus;
        cb.LostKeyboardFocus -= hook.LostKeyboardFocus;
        cb.Unloaded -= ComboBoxHook_Unloaded;
        cb.Tag = null;
    }
}
