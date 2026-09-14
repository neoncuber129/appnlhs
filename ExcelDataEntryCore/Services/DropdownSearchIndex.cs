namespace ExcelDataEntryApp.Services;

/// <summary>Danh sách dropdown đã sắp xếp, tìm kiếm một lần duyệt không cấp phát LINQ.</summary>
public sealed class DropdownSearchIndex
{
    private readonly string[] _options;

    public DropdownSearchIndex(IEnumerable<string> options)
    {
        _options = options
            .Where(static v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static v => v, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    public int Count => _options.Length;

    public IReadOnlyList<string> AllOptions => _options;

    public IReadOnlyList<string> Search(string filter, int maxItems, string? ensureValue)
    {
        var limit = maxItems <= 0 ? int.MaxValue : maxItems;
        var trimmedFilter = filter.Trim();

        if (trimmedFilter.Length == 0)
        {
            return EnsureValueInResults(_options, ensureValue, limit);
        }

        var comparison = StringComparison.CurrentCultureIgnoreCase;
        var results = new List<string>(Math.Min(limit, 64));
        for (var i = 0; i < _options.Length && results.Count < limit; i++)
        {
            if (_options[i].Contains(trimmedFilter, comparison))
            {
                results.Add(_options[i]);
            }
        }

        return EnsureValueInResults(results, ensureValue, limit);
    }

    private static IReadOnlyList<string> EnsureValueInResults(
        IReadOnlyList<string> results,
        string? ensureValue,
        int limit)
    {
        if (string.IsNullOrWhiteSpace(ensureValue))
        {
            return results;
        }

        foreach (var item in results)
        {
            if (string.Equals(item, ensureValue, StringComparison.Ordinal))
            {
                return results;
            }
        }

        if (results is List<string> mutable)
        {
            mutable.Insert(0, ensureValue);
            if (mutable.Count > limit)
            {
                mutable.RemoveAt(mutable.Count - 1);
            }

            return mutable;
        }

        var copy = new List<string>(results.Count + 1) { ensureValue };
        copy.AddRange(results);
        if (copy.Count > limit)
        {
            copy.RemoveAt(copy.Count - 1);
        }

        return copy;
    }
}
