namespace ExcelDataEntryApp.Models;

public static class DropdownLimits
{
    /// <summary>Trên ngưỡng này dùng ComboBox có lọc thay vì bind toàn bộ danh sách.</summary>
    public const int SearchableThreshold = 80;

    /// <summary>0 = không giới hạn số mục hiển thị (dùng virtualization).</summary>
    public const int UnlimitedVisibleItems = 0;

    /// <summary>Từ ngưỡng này tìm kiếm / nạp danh sách chạy nền để không block UI.</summary>
    public const int BackgroundSearchThreshold = 1500;

    /// <summary>Độ trễ gõ phím trước khi lọc (ms).</summary>
    public const int SearchDebounceMilliseconds = 220;
}
