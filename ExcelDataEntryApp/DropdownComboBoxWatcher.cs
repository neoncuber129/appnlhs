using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp;

/// <summary>
/// ComboBox dropdown lớn: lọc có debounce, không tự bung list khi focus.
/// </summary>
public static class DropdownComboBoxWatcher
{
    public static readonly DependencyProperty EnableProperty = DependencyProperty.RegisterAttached(
        "Enable",
        typeof(bool),
        typeof(DropdownComboBoxWatcher),
        new PropertyMetadata(false, OnEnableChanged));

    public static void SetEnable(ComboBox element, bool value) => element.SetValue(EnableProperty, value);

    public static bool GetEnable(ComboBox element) => (bool)element.GetValue(EnableProperty);

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

            if (IsClickInsideComboBoxScope(cb, originalSource) || IsOriginalSourceUnderComboPopup(cb, originalSource))
            {
                continue;
            }

            cb.IsDropDownOpen = false;
        }
    }

    private sealed class Hook
    {
        public CaretRestoreGate CaretGate { get; } = new();
        public TextChangedEventHandler TextChanged { get; set; } = null!;
        public EventHandler DropDownOpened { get; set; } = null!;
        public EventHandler DropDownClosed { get; set; } = null!;
        public KeyboardFocusChangedEventHandler LostKeyboardFocus { get; set; } = null!;
        public DispatcherTimer DebounceTimer { get; } = new()
        {
            Interval = TimeSpan.FromMilliseconds(DropdownLimits.SearchDebounceMilliseconds)
        };

        public bool OpenDropdownAfterSearch { get; set; }
        public bool SuppressSearch { get; set; }
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

        var hook = new Hook();
        hook.TextChanged = (_, _) => OnDropdownTextChanged(cb, hook);
        hook.DropDownOpened = (_, _) => OnDropDownOpened(cb, hook);
        hook.DropDownClosed = (_, _) => hook.SuppressSearch = false;
        hook.LostKeyboardFocus = (_, _) =>
        {
            cb.IsDropDownOpen = false;
            if (cb.DataContext is EditableField field)
            {
                if (cb.IsEditable)
                {
                    field.CommitValueFromEditor(cb.Text?.Trim() ?? string.Empty);
                }

                field.ValidateDropdownValue();
            }
        };
        hook.DebounceTimer.Tick += (_, _) =>
        {
            hook.DebounceTimer.Stop();
            RunSearch(cb, hook, hook.OpenDropdownAfterSearch);
        };

        cb.Tag = hook;
        TextEditorCaretHelper.AttachCaretRestoreGate(cb, hook.CaretGate);
        cb.AddHandler(TextBoxBase.TextChangedEvent, hook.TextChanged);
        cb.DropDownOpened += hook.DropDownOpened;
        cb.DropDownClosed += hook.DropDownClosed;
        cb.LostKeyboardFocus += hook.LostKeyboardFocus;
        cb.Unloaded += ComboBoxHook_Unloaded;
    }

    private static void OnDropDownOpened(ComboBox cb, Hook hook)
    {
        hook.DebounceTimer.Stop();
        hook.OpenDropdownAfterSearch = true;
        RunSearch(cb, hook, openDropdownIfResults: true);
    }

    private static void OnDropdownTextChanged(ComboBox cb, Hook hook)
    {
        if (cb.DataContext is EditableField field)
        {
            var text = cb.Text ?? string.Empty;
            _ = TextEditorCaretHelper.TryGetEditableTextBox(cb, out var editableTextBox);
            var caret = editableTextBox?.CaretIndex ?? text.Length;
            var selectionLength = editableTextBox?.SelectionLength ?? 0;
            var token = hook.CaretGate.Next();
            field.SetValueSilently(text);
            if (editableTextBox is not null)
            {
                TextEditorCaretHelper.RestoreTextBoxCaret(editableTextBox, caret, selectionLength, hook.CaretGate, token);
            }
        }

        if (hook.SuppressSearch)
        {
            return;
        }

        hook.OpenDropdownAfterSearch = cb.IsDropDownOpen;
        hook.DebounceTimer.Stop();
        hook.DebounceTimer.Start();
    }

    private static void RunSearch(ComboBox cb, Hook hook, bool openDropdownIfResults)
    {
        if (cb.DataContext is not EditableField field)
        {
            return;
        }

        field.SearchDropdownOptions(
            cb.Text,
            field.Value,
            results =>
            {
                if (!cb.IsLoaded)
                {
                    return;
                }

                TextEditorCaretHelper.PreserveEditableComboCaret(
                    cb,
                    hook.CaretGate,
                    () => field.ApplyFilteredResults(results));

                if (!cb.IsKeyboardFocusWithin)
                {
                    return;
                }

                if (results.Count == 0)
                {
                    cb.IsDropDownOpen = false;
                    return;
                }

                if (openDropdownIfResults)
                {
                    cb.IsDropDownOpen = true;
                }
            });
    }

    private static void ComboBoxHook_Unloaded(object sender, RoutedEventArgs e) => Unhook(sender as ComboBox);

    private static void Unhook(ComboBox? cb)
    {
        if (cb?.Tag is not Hook hook)
        {
            return;
        }

        hook.DebounceTimer.Stop();
        cb.RemoveHandler(TextBoxBase.TextChangedEvent, hook.TextChanged);
        cb.DropDownOpened -= hook.DropDownOpened;
        cb.DropDownClosed -= hook.DropDownClosed;
        cb.LostKeyboardFocus -= hook.LostKeyboardFocus;
        cb.Unloaded -= ComboBoxHook_Unloaded;
        cb.Tag = null;
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
}
