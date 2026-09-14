using ExcelDataEntryApp.Infrastructure;
using ExcelDataEntryApp.Services;

namespace ExcelDataEntryApp.Models;

public sealed class EditableField : ObservableObject
{
    private string _value = string.Empty;
    private IReadOnlyList<string> _dropdownOptions = [];
    private IReadOnlyList<string> _suggestionOptions = [];
    private IReadOnlyList<string> _displayDropdownItems = [];
    private DropdownSearchIndex? _dropdownSearchIndex;
    private HashSet<string> _dropdownOptionLookup = new(StringComparer.OrdinalIgnoreCase);
    private int _searchGeneration;
    private bool _isDropdownValueInvalid;
    private string _dropdownValidationMessage = string.Empty;
    private bool _usesDropdownEditor;
    private bool _suppressBindingNotification;

    public required string HeaderName { get; init; }

    /// <summary>Nhãn hiển thị (thường là phần con khi đã gom nhóm header cha).</summary>
    public required string HeaderDisplayName { get; init; }

    public string ParentHeaderName { get; init; } = string.Empty;

    public bool ShowParentHeader { get; init; }

    public bool IsGroupedUnderParentHeader { get; init; }

    public bool IsFirstInHeaderGroup { get; init; }

    public bool IsLastInHeaderGroup { get; init; }

    public string GroupBorderHex { get; init; } = "Transparent";

    public bool HasGroupBorder =>
        IsGroupedUnderParentHeader
        && !string.Equals(GroupBorderHex, "Transparent", StringComparison.OrdinalIgnoreCase);

    public string ParentTitleBackgroundHex { get; init; } = "Transparent";

    public string ParentTitleForegroundHex { get; init; } = "#111827";

    public bool HasParentTitleHighlight =>
        ShowParentHeader
        && !string.Equals(ParentTitleBackgroundHex, "Transparent", StringComparison.OrdinalIgnoreCase);

    public required int ColumnIndex { get; init; }

    /// <summary>Sheet chứa ô (chế độ ghép nhiều sheet).</summary>
    public string SheetName { get; init; } = string.Empty;

    /// <summary>Hiện tiêu đề ngăn cách nhóm sheet (dòng đầu mỗi sheet).</summary>
    public bool ShowSheetSeparator { get; init; }

    public string SheetSeparatorTitle { get; init; } = string.Empty;

    /// <summary>Hàng Excel thực khi load field.</summary>
    public int SourceRowIndex { get; init; }

    public IReadOnlyList<string> DropdownOptions => _dropdownOptions;

    public IReadOnlyList<int> ParentDropdownColumns { get; init; } = [];

    public bool IsDependentDropdown => ParentDropdownColumns.Count > 0;

    public IReadOnlyList<string> DisplayDropdownItems => _displayDropdownItems;

    public bool HasDropdown => _dropdownOptions.Count > 0;

      /// <summary>Luôn giữ ComboBox sau khi đã xác định là cột dropdown (kể cả danh sách tạm rỗng).</summary>
    public bool ShowDropdownEditor => _usesDropdownEditor || HasDropdown || IsDependentDropdown;

    public bool HasLargeDropdown { get; private set; }

    public bool HasSmallDropdown => ShowDropdownEditor && !HasLargeDropdown;

    public bool IsDropdownValueInvalid
    {
        get => _isDropdownValueInvalid;
        private set => SetProperty(ref _isDropdownValueInvalid, value);
    }

    public string DropdownValidationMessage
    {
        get => _dropdownValidationMessage;
        private set => SetProperty(ref _dropdownValidationMessage, value);
    }

    public void ConfigureDropdownRole(bool isDependentColumn, IReadOnlyList<string> initialOptions)
    {
        if (isDependentColumn || initialOptions.Any(static value => !string.IsNullOrWhiteSpace(value)))
        {
            _usesDropdownEditor = true;
            OnPropertyChanged(nameof(ShowDropdownEditor));
            OnPropertyChanged(nameof(HasSuggestions));
            OnPropertyChanged(nameof(HasSmallDropdown));
        }
    }

    public void SetInitialDropdownOptions(IReadOnlyList<string> options)
    {
        if ((IsDependentDropdown || _usesDropdownEditor) && options.Count == 0)
        {
            options = [string.Empty];
        }

        if (options.Any(static value => !string.IsNullOrWhiteSpace(value)))
        {
            _usesDropdownEditor = true;
        }

        _dropdownOptions = options;
        OnPropertyChanged(nameof(DropdownOptions));
        OnPropertyChanged(nameof(HasDropdown));
        OnPropertyChanged(nameof(ShowDropdownEditor));
        OnPropertyChanged(nameof(HasSuggestions));
        OnPropertyChanged(nameof(HasSmallDropdown));
    }

    public void ReplaceDropdownOptions(IReadOnlyList<string> options, DropdownSearchIndex? searchIndex = null)
    {
        if (options.Any(static value => !string.IsNullOrWhiteSpace(value)))
        {
            _usesDropdownEditor = true;
            SuggestionOptions = [];
        }

        if ((IsDependentDropdown || _usesDropdownEditor) && options.Count == 0)
        {
            options = [string.Empty];
        }

        _dropdownOptions = options;

        OnPropertyChanged(nameof(DropdownOptions));
        OnPropertyChanged(nameof(HasDropdown));
        OnPropertyChanged(nameof(ShowDropdownEditor));
        OnPropertyChanged(nameof(HasSuggestions));
        InitializeDropdownPresentation(searchIndex);
        NotifyDropdownOptionsRefreshed();
    }

    public void InitializeDropdownPresentation(DropdownSearchIndex? searchIndex = null)
    {
        _dropdownSearchIndex = searchIndex ?? new DropdownSearchIndex(_dropdownOptions);
        _dropdownOptionLookup = new HashSet<string>(
            _dropdownOptions.Where(static v => v is not null),
            StringComparer.OrdinalIgnoreCase);
        HasLargeDropdown = ShowDropdownEditor
            && !IsDependentDropdown
            && _dropdownSearchIndex.Count > DropdownLimits.SearchableThreshold;
        OnPropertyChanged(nameof(HasLargeDropdown));
        OnPropertyChanged(nameof(HasSmallDropdown));
        OnPropertyChanged(nameof(ShowDropdownEditor));
        OnPropertyChanged(nameof(HasSuggestions));

        if (HasLargeDropdown)
        {
            SearchDropdownOptions(string.Empty, Value, ApplyFilteredResults);
        }
        else
        {
            _displayDropdownItems = [];
            OnPropertyChanged(nameof(DisplayDropdownItems));
        }

        ValidateDropdownValue();
    }

    public bool IsDropdownValueAllowed(string? value)
    {
        if (!ShowDropdownEditor || string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        return _dropdownOptionLookup.Contains(value.Trim());
    }

    public bool HasNonEmptyDropdownValue =>
        ShowDropdownEditor && !string.IsNullOrWhiteSpace(Value);

    public void ValidateDropdownValue()
    {
        if (!ShowDropdownEditor)
        {
            ClearDropdownValidation();
            return;
        }

        var effectiveValue = (Value ?? string.Empty).Trim();
        if (effectiveValue.Length == 0)
        {
            AssignValue(string.Empty, notifyBinding: ShouldNotifyBinding());
            ClearDropdownValidation();
            return;
        }

        if (IsDropdownValueAllowed(effectiveValue))
        {
            AssignValue(effectiveValue, notifyBinding: ShouldNotifyBinding());
            ClearDropdownValidation();
            return;
        }

        // ComboBox không chỉnh sửa không hiển thị giá trị lạ — coi như để trống.
        if (HasSmallDropdown)
        {
            AssignValue(string.Empty, notifyBinding: ShouldNotifyBinding());
            ClearDropdownValidation();
            return;
        }

        IsDropdownValueInvalid = true;
        DropdownValidationMessage = $"Giá trị \"{Value}\" không có trong danh sách dropdown.";
    }

    private void ClearDropdownValidation()
    {
        IsDropdownValueInvalid = false;
        DropdownValidationMessage = string.Empty;
    }

    public void SearchDropdownOptions(string? filter, string? ensureValue, Action<IReadOnlyList<string>> applyResults)
    {
        if (_dropdownSearchIndex is null)
        {
            applyResults([]);
            return;
        }

        var generation = Interlocked.Increment(ref _searchGeneration);
        var trimmedFilter = filter?.Trim() ?? string.Empty;
        var runSearch = () => _dropdownSearchIndex.Search(
            trimmedFilter,
            DropdownLimits.UnlimitedVisibleItems,
            ensureValue);

        if (_dropdownSearchIndex.Count < DropdownLimits.BackgroundSearchThreshold)
        {
            applyResults(runSearch());
            return;
        }

        Task.Run(runSearch)
            .ContinueWith(
                task =>
                {
                    if (generation != Volatile.Read(ref _searchGeneration))
                    {
                        return;
                    }

                    applyResults(task.Result);
                },
                CancellationToken.None,
                TaskContinuationOptions.None,
                TaskScheduler.FromCurrentSynchronizationContext());
    }

    public void ApplyFilteredResults(IReadOnlyList<string> items)
    {
        if (ReferenceEquals(_displayDropdownItems, items)
            || SequencesEqual(_displayDropdownItems, items))
        {
            return;
        }

        _displayDropdownItems = items;
        OnPropertyChanged(nameof(DisplayDropdownItems));
    }

    private static bool SequencesEqual(IReadOnlyList<string> current, IReadOnlyList<string> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            if (!string.Equals(current[i], next[i], StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Row strip / panel background for dropdown fields (hex or Transparent).</summary>
    public string RowHighlightBackgroundHex { get; init; } = "Transparent";

    public bool HasRowHighlight =>
        !string.Equals(RowHighlightBackgroundHex, "Transparent", StringComparison.OrdinalIgnoreCase);

    /// <summary>Border color for the highlighted row (hex or Transparent).</summary>
    public string RowHighlightBorderHex { get; init; } = "Transparent";

    /// <summary>Header label foreground on this row (hex).</summary>
    public string RowHeaderForegroundHex { get; init; } = "#111827";

    /// <summary>ComboBox surface on dropdown rows (hex).</summary>
    public string RowInputBackgroundHex { get; init; } = "#FFFFFF";

    /// <summary>ComboBox text on dropdown rows (hex).</summary>
    public string RowInputForegroundHex { get; init; } = "#111827";

    public IReadOnlyList<string> SuggestionOptions
    {
        get => _suggestionOptions;
        set
        {
            if (SetProperty(ref _suggestionOptions, value))
            {
                OnPropertyChanged(nameof(HasSuggestions));
            }
        }
    }

    public bool HasSuggestions => SuggestionOptions.Count > 0 && !ShowDropdownEditor;

    public string Value
    {
        get => _value;
        set
        {
            if (!AssignValue(value ?? string.Empty, notifyBinding: true))
            {
                return;
            }

            ValidateDropdownValue();
            ValueChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>Cập nhật giá trị khi đang gõ — không đẩy ngược lên TextBox (tránh nhảy con trỏ).</summary>
    public void SetValueSilently(string value)
    {
        _suppressBindingNotification = true;
        try
        {
            if (!AssignValue(value ?? string.Empty, notifyBinding: false))
            {
                return;
            }

            ValueChanged?.Invoke(this, EventArgs.Empty);
        }
        finally
        {
            _suppressBindingNotification = false;
        }
    }

    /// <summary>Đồng bộ binding khi rời ô nhập.</summary>
    public void CommitValueFromEditor(string value)
    {
        var normalized = value ?? string.Empty;
        if (string.Equals(_value, normalized, StringComparison.Ordinal))
        {
            return;
        }

        if (SetProperty(ref _value, normalized))
        {
            ValidateDropdownValue();
            ValueChanged?.Invoke(this, EventArgs.Empty);
            return;
        }

        ValidateDropdownValue();
    }

    private bool AssignValue(string value, bool notifyBinding)
    {
        if (string.Equals(_value, value, StringComparison.Ordinal))
        {
            return false;
        }

        _value = value;
        if (notifyBinding && ShouldNotifyBinding())
        {
            OnPropertyChanged(nameof(Value));
        }

        return true;
    }

    private bool ShouldNotifyBinding() => !_suppressBindingNotification;

    public event EventHandler? ValueChanged;

    public event EventHandler? DropdownSelectionCommitted;

    public void NotifyDropdownSelectionCommitted() =>
        DropdownSelectionCommitted?.Invoke(this, EventArgs.Empty);

    public int DropdownOptionsRevision { get; private set; }

    public int OpenDropdownRequestId { get; private set; }

    public void NotifyDropdownOptionsRefreshed()
    {
        DropdownOptionsRevision++;
        OnPropertyChanged(nameof(DropdownOptionsRevision));
    }

    public void RequestOpenDropdown()
    {
        OpenDropdownRequestId++;
        OnPropertyChanged(nameof(OpenDropdownRequestId));
    }
}
