using System.Globalization;
using System.Text;

namespace ExcelDataEntryApp.Infrastructure;

public static class VietnameseTextHelper
{
    private static readonly HashSet<char> ToneCombiningMarks =
    [
        '\u0300', // huyền
        '\u0301', // sắc
        '\u0303', // ngã
        '\u0309', // hỏi
        '\u0323', // nặng
    ];

    /// <summary>
    /// Chuẩn hóa key so khớp liên kết số liệu: bỏ vị trí dấu thanh trên vần, giữ loại dấu và đ/Đ.
    /// Ví dụ: hòa và hoà cùng key; hóa khác hòa; Nguyen khác Nguyễn.
    /// </summary>
    public static string ToTonePlacementKey(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var trimmed = text.Trim();
        var decomposed = trimmed.Normalize(NormalizationForm.FormD);
        var baseBuilder = new StringBuilder(decomposed.Length);
        var toneBuilder = new StringBuilder(decomposed.Length);

        foreach (var ch in decomposed)
        {
            if (ToneCombiningMarks.Contains(ch))
            {
                toneBuilder.Append(ClassifyToneMark(ch).ToString(CultureInfo.InvariantCulture));
                continue;
            }

            if (CharUnicodeInfo.GetUnicodeCategory(ch) == UnicodeCategory.NonSpacingMark)
            {
                baseBuilder.Append(ch);
                continue;
            }

            baseBuilder.Append(ch);
        }

        var baseLetters = baseBuilder.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
        if (toneBuilder.Length == 0)
        {
            return baseLetters + '0';
        }

        return baseLetters + toneBuilder;
    }

    public static string ToSearchKey(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var normalized = text.Trim()
            .Replace('đ', 'd')
            .Replace('Đ', 'D')
            .Normalize(NormalizationForm.FormD);

        var builder = new StringBuilder(normalized.Length);
        foreach (var ch in normalized)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
            {
                builder.Append(ch);
            }
        }

        return builder.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
    }

    public static bool ContainsNormalized(string? haystack, string? needle)
    {
        if (string.IsNullOrWhiteSpace(needle))
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(haystack))
        {
            return false;
        }

        return ToSearchKey(haystack).Contains(ToSearchKey(needle), StringComparison.Ordinal);
    }

    private static int ClassifyToneMark(char mark) => mark switch
    {
        '\u0301' => 1, // sắc
        '\u0300' => 2, // huyền
        '\u0309' => 3, // hỏi
        '\u0303' => 4, // ngã
        '\u0323' => 5, // nặng
        _ => 0
    };
}
