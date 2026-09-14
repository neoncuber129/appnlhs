using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp;

/// <summary>
/// Đồng bộ giá trị dropdown ngay khi chọn (không đợi lost focus), báo refresh dropdown con,
/// và tự bung dropdown khi được yêu cầu (sau khi dropdown cha thay đổi).
/// </summary>
public static class DependentDropdownComboBoxWatcher
{
    public static readonly DependencyProperty EnableProperty = DependencyProperty.RegisterAttached(
        "Enable",
        typeof(bool),
        typeof(DependentDropdownComboBoxWatcher),
        new PropertyMetadata(false, OnEnableChanged));

    public static void SetEnable(ComboBox element, bool value) => element.SetValue(EnableProperty, value);

    public static bool GetEnable(ComboBox element) => (bool)element.GetValue(EnableProperty);

    private static readonly DependencyProperty HookProperty = DependencyProperty.RegisterAttached(
        "Hook",
        typeof(Hook),
        typeof(DependentDropdownComboBoxWatcher));

    private sealed class Hook
    {
        public SelectionChangedEventHandler SelectionChanged { get; set; } = null!;
        public EventHandler DropDownClosed { get; set; } = null!;
        public KeyboardFocusChangedEventHandler LostKeyboardFocus { get; set; } = null!;
        public DependencyPropertyChangedEventHandler DataContextChanged { get; set; } = null!;
        public PropertyChangedEventHandler? FieldPropertyChanged { get; set; }
        public EditableField? SubscribedField { get; set; }
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
        if (sender is not ComboBox cb)
        {
            return;
        }

        if (cb.GetValue(HookProperty) is not Hook hook)
        {
            hook = new Hook();
            hook.SelectionChanged = (_, args) => OnSelectionChanged(cb, args);
            hook.DropDownClosed = (_, _) => OnDropDownClosed(cb);
            hook.LostKeyboardFocus = (_, _) => RestoreSelectionFromField(cb);
            hook.DataContextChanged = (_, _) => SubscribeField(cb, hook);

            cb.SetValue(HookProperty, hook);
            cb.SelectionChanged += hook.SelectionChanged;
            cb.DropDownClosed += hook.DropDownClosed;
            cb.LostKeyboardFocus += hook.LostKeyboardFocus;
            cb.DataContextChanged += hook.DataContextChanged;
            cb.Unloaded += ComboBoxHook_Unloaded;
        }

        SubscribeField(cb, hook);
        RestoreSelectionFromField(cb);
    }

    private static void SubscribeField(ComboBox cb, Hook hook)
    {
        UnsubscribeField(hook);
        if (cb.DataContext is not EditableField field)
        {
            return;
        }

        hook.SubscribedField = field;
        hook.FieldPropertyChanged = (_, args) =>
        {
            if (args.PropertyName == nameof(EditableField.OpenDropdownRequestId))
            {
                ScheduleOpenDropdown(cb);
            }
            else if (args.PropertyName is nameof(EditableField.Value) or nameof(EditableField.DropdownOptionsRevision))
            {
                RestoreSelectionFromField(cb);
            }
        };
        field.PropertyChanged += hook.FieldPropertyChanged;
    }

    private static void UnsubscribeField(Hook hook)
    {
        if (hook.SubscribedField is not null && hook.FieldPropertyChanged is not null)
        {
            hook.SubscribedField.PropertyChanged -= hook.FieldPropertyChanged;
        }

        hook.SubscribedField = null;
        hook.FieldPropertyChanged = null;
    }

    private static void ScheduleOpenDropdown(ComboBox cb)
    {
        cb.Dispatcher.BeginInvoke(
            () => TryOpenDropdown(cb),
            DispatcherPriority.ApplicationIdle);
    }

    private static void TryOpenDropdown(ComboBox cb)
    {
        if (!cb.IsLoaded || !cb.IsVisible || !cb.IsEnabled)
        {
            return;
        }

        if (cb.DataContext is not EditableField field || !field.ShowDropdownEditor)
        {
            return;
        }

        if (!field.DropdownOptions.Any(static value => !string.IsNullOrWhiteSpace(value)))
        {
            return;
        }

        RestoreSelectionFromField(cb);
        cb.Focus();
        Keyboard.Focus(cb);
        cb.IsDropDownOpen = true;
    }

    private static void OnSelectionChanged(ComboBox cb, SelectionChangedEventArgs e)
    {
        if (cb.DataContext is not EditableField field || e.AddedItems.Count == 0)
        {
            return;
        }

        var selected = e.AddedItems[0] as string ?? Convert.ToString(e.AddedItems[0]) ?? string.Empty;
        CommitFieldValue(field, selected);
        field.NotifyDropdownSelectionCommitted();
    }

    private static void OnDropDownClosed(ComboBox cb)
    {
        if (cb.DataContext is not EditableField field)
        {
            return;
        }

        if (cb.IsEditable)
        {
            CommitFieldValue(field, cb.Text?.Trim() ?? string.Empty);
            field.NotifyDropdownSelectionCommitted();
            return;
        }

        if (cb.SelectedItem is string selected && !string.IsNullOrEmpty(selected))
        {
            CommitFieldValue(field, selected);
            field.NotifyDropdownSelectionCommitted();
            return;
        }

        RestoreSelectionFromField(cb);
    }

    private static void RestoreSelectionFromField(ComboBox cb)
    {
        if (cb.DataContext is not EditableField field || cb.IsEditable)
        {
            return;
        }

        field.ValidateDropdownValue();

        var value = field.Value ?? string.Empty;
        if (string.Equals(cb.SelectedItem as string, value, StringComparison.Ordinal))
        {
            return;
        }

        if (string.IsNullOrEmpty(value))
        {
            if (cb.SelectedItem is not null)
            {
                cb.SelectedItem = null;
            }

            return;
        }

        if (field.DropdownOptions.Contains(value))
        {
            cb.SelectedItem = value;
        }
        else if (cb.SelectedItem is not null)
        {
            cb.SelectedItem = null;
        }
    }

    private static void CommitFieldValue(EditableField field, string value)
    {
        if (!string.Equals(field.Value, value, StringComparison.Ordinal))
        {
            field.Value = value;
        }
    }

    private static void ComboBoxHook_Unloaded(object sender, RoutedEventArgs e) => Unhook(sender as ComboBox);

    private static void Unhook(ComboBox? cb)
    {
        if (cb?.GetValue(HookProperty) is not Hook hook)
        {
            return;
        }

        UnsubscribeField(hook);
        cb.SelectionChanged -= hook.SelectionChanged;
        cb.DropDownClosed -= hook.DropDownClosed;
        cb.LostKeyboardFocus -= hook.LostKeyboardFocus;
        cb.DataContextChanged -= hook.DataContextChanged;
        cb.Unloaded -= ComboBoxHook_Unloaded;
        cb.ClearValue(HookProperty);
    }
}
