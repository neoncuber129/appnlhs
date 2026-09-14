using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using ExcelDataEntryApp.Models;

namespace ExcelDataEntryApp;

internal static class ImportListColumnFilterPopup
{
    private const double RowHeight = 28;
    private const int ShowAllItemThreshold = 24;

    public static void Show(
        UIElement anchor,
        ImportListColumnFilterState filterState,
        IReadOnlyList<string> distinctValues,
        Action<ImportListColumnFilterState> onApplied)
    {
        var selected = filterState.AllowedValues is null
            ? distinctValues.ToHashSet(StringComparer.Ordinal)
            : new HashSet<string>(filterState.AllowedValues, StringComparer.Ordinal);

        var checkBoxes = new List<CheckBox>();
        var listPanel = new StackPanel();

        void RebuildList(string? query)
        {
            listPanel.Children.Clear();
            checkBoxes.Clear();
            foreach (var value in distinctValues)
            {
                if (!string.IsNullOrWhiteSpace(query)
                    && !value.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var checkBox = new CheckBox
                {
                    Content = string.IsNullOrEmpty(value) ? "(Trống)" : value,
                    IsChecked = selected.Contains(value),
                    Margin = new Thickness(0, 2, 0, 2)
                };
                checkBox.Checked += (_, _) => selected.Add(value);
                checkBox.Unchecked += (_, _) => selected.Remove(value);
                checkBoxes.Add(checkBox);
                listPanel.Children.Add(checkBox);
            }
        }

        RebuildList(null);

        var searchBox = new TextBox
        {
            MinHeight = 28,
            Margin = new Thickness(0, 0, 0, 6),
            ToolTip = "Tìm trong danh sách giá trị cột"
        };

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = listPanel
        };

        void UpdateListSize(int visibleCount)
        {
            var workHeight = SystemParameters.WorkArea.Height;
            var maxListHeight = Math.Max(280, workHeight * 0.55);
            var idealHeight = Math.Max(40, visibleCount * RowHeight + 8);

            if (visibleCount <= ShowAllItemThreshold && idealHeight <= maxListHeight)
            {
                scroll.MaxHeight = double.PositiveInfinity;
                scroll.Height = idealHeight;
                scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            }
            else
            {
                scroll.Height = double.NaN;
                scroll.MaxHeight = maxListHeight;
                scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
        }

        UpdateListSize(distinctValues.Count);

        searchBox.TextChanged += (_, _) =>
        {
            RebuildList(searchBox.Text.Trim());
            UpdateListSize(checkBoxes.Count);
        };

        var selectAllButton = new Button { Content = "Chọn tất cả", Padding = new Thickness(8, 2, 8, 2), Margin = new Thickness(0, 0, 6, 0) };
        selectAllButton.Click += (_, _) =>
        {
            selected.Clear();
            foreach (var value in distinctValues)
            {
                selected.Add(value);
            }

            foreach (var checkBox in checkBoxes)
            {
                checkBox.IsChecked = true;
            }
        };

        var clearButton = new Button { Content = "Bỏ chọn", Padding = new Thickness(8, 2, 8, 2) };
        clearButton.Click += (_, _) =>
        {
            selected.Clear();
            foreach (var checkBox in checkBoxes)
            {
                checkBox.IsChecked = false;
            }
        };

        var buttonRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 6) };
        buttonRow.Children.Add(selectAllButton);
        buttonRow.Children.Add(clearButton);

        var applyButton = new Button
        {
            Content = "Áp dụng",
            Padding = new Thickness(10, 4, 10, 4),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var root = new Border
        {
            Padding = new Thickness(12),
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = System.Windows.Media.Brushes.LightGray,
            BorderThickness = new Thickness(1),
            Child = new StackPanel()
        };
        root.MinWidth = ComputePopupMinWidth(distinctValues);
        root.MaxWidth = Math.Min(520, SystemParameters.WorkArea.Width * 0.45);

        var stack = (StackPanel)root.Child;
        stack.Children.Add(new TextBlock
        {
            Text = filterState.HeaderName,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 6),
            TextWrapping = TextWrapping.Wrap
        });
        stack.Children.Add(searchBox);
        stack.Children.Add(buttonRow);
        stack.Children.Add(scroll);
        stack.Children.Add(applyButton);

        var popup = new Popup
        {
            PlacementTarget = anchor,
            Placement = PlacementMode.Bottom,
            StaysOpen = false,
            AllowsTransparency = true,
            Child = root
        };

        applyButton.Click += (_, _) =>
        {
            if (selected.Count == 0)
            {
                filterState.AllowedValues = [];
            }
            else if (selected.Count == distinctValues.Count)
            {
                filterState.AllowedValues = null;
            }
            else
            {
                filterState.AllowedValues = new HashSet<string>(selected, StringComparer.Ordinal);
            }

            onApplied(filterState);
            popup.IsOpen = false;
        };

        popup.IsOpen = true;
    }

    private static double ComputePopupMinWidth(IReadOnlyList<string> distinctValues)
    {
        var longest = distinctValues
            .Select(value => string.IsNullOrEmpty(value) ? "(Trống)" : value)
            .DefaultIfEmpty(string.Empty)
            .MaxBy(value => value.Length)
            ?.Length ?? 0;

        return Math.Clamp(280 + longest * 5.5, 300, 480);
    }
}
