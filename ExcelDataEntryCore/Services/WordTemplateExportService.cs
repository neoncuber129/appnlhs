using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelDataEntryApp.Services;

public static class WordTemplateExportService
{
    private static readonly Regex PlaceholderRegex = new(@"\$(\d+|[A-Za-z_][A-Za-z0-9_]*)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static void ExportFromTemplate(
        string templatePath,
        string outputPath,
        IReadOnlyDictionary<int, string> rowValues,
        IReadOnlyDictionary<string, string>? namedPlaceholders = null)
    {
        File.Copy(templatePath, outputPath, overwrite: true);
        using var doc = WordprocessingDocument.Open(outputPath, true);

        ReplaceInPart(doc.MainDocumentPart, rowValues, namedPlaceholders);
        if (doc.MainDocumentPart?.HeaderParts is not null)
        {
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                ReplaceInPart(header, rowValues, namedPlaceholders);
            }
        }

        if (doc.MainDocumentPart?.FooterParts is not null)
        {
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                ReplaceInPart(footer, rowValues, namedPlaceholders);
            }
        }
    }

    public static void ExportManyRecordsToSingleFile(
        string templatePath,
        string outputPath,
        IReadOnlyList<IReadOnlyDictionary<int, string>> records,
        IReadOnlyList<IReadOnlyDictionary<string, string>>? namedPlaceholderRecords = null,
        Action<int, int>? progressCallback = null)
    {
        if (records.Count == 0)
        {
            throw new ArgumentException("Không có dữ liệu để xuất.", nameof(records));
        }

        var firstNamed = namedPlaceholderRecords is not null && namedPlaceholderRecords.Count > 0
            ? namedPlaceholderRecords[0]
            : null;
        ExportFromTemplate(templatePath, outputPath, records[0], firstNamed);
        progressCallback?.Invoke(1, records.Count);
        for (var i = 1; i < records.Count; i++)
        {
            var tempFile = Path.Combine(Path.GetTempPath(), $"hssk_merge_{Guid.NewGuid():N}.docx");
            try
            {
                var named = namedPlaceholderRecords is not null && i < namedPlaceholderRecords.Count
                    ? namedPlaceholderRecords[i]
                    : null;
                ExportFromTemplate(templatePath, tempFile, records[i], named);
                AppendDocumentWithPageBreak(outputPath, tempFile);
                progressCallback?.Invoke(i + 1, records.Count);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }

    private static void ReplaceInPart(
        OpenXmlPart? part,
        IReadOnlyDictionary<int, string> rowValues,
        IReadOnlyDictionary<string, string>? namedPlaceholders)
    {
        if (part?.RootElement is null)
        {
            return;
        }

        foreach (var paragraph in part.RootElement.Descendants<Paragraph>())
        {
            var textNodes = paragraph.Descendants<Text>().ToList();
            if (textNodes.Count == 0)
            {
                continue;
            }

            var merged = string.Concat(textNodes.Select(t => t.Text));
            if (string.IsNullOrEmpty(merged))
            {
                continue;
            }

            if (!PlaceholderRegex.IsMatch(merged))
            {
                continue;
            }

            ReplacePlaceholderAcrossTextNodes(textNodes, rowValues, namedPlaceholders);
        }
    }

    private static void ReplacePlaceholderAcrossTextNodes(
        IReadOnlyList<Text> textNodes,
        IReadOnlyDictionary<int, string> rowValues,
        IReadOnlyDictionary<string, string>? namedPlaceholders)
    {
        if (textNodes.Count == 0)
        {
            return;
        }

        var fullText = string.Concat(textNodes.Select(t => t.Text));
        if (string.IsNullOrEmpty(fullText))
        {
            return;
        }

        var matches = PlaceholderRegex.Matches(fullText);
        if (matches.Count == 0)
        {
            return;
        }

        var starts = new int[textNodes.Count];
        var lengths = new int[textNodes.Count];
        var cursor = 0;
        for (var i = 0; i < textNodes.Count; i++)
        {
            starts[i] = cursor;
            var len = textNodes[i].Text?.Length ?? 0;
            lengths[i] = len;
            cursor += len;
        }

        for (var m = matches.Count - 1; m >= 0; m--)
        {
            var match = matches[m];
            var token = match.Groups[1].Value;
            string replacement;
            if (int.TryParse(token, out var col))
            {
                replacement = rowValues.TryGetValue(col, out var value) ? value : string.Empty;
            }
            else if (namedPlaceholders is not null && namedPlaceholders.TryGetValue(token, out var namedValue))
            {
                replacement = namedValue ?? string.Empty;
            }
            else
            {
                continue;
            }

            ReplaceRangeInTextNodes(textNodes, starts, lengths, match.Index, match.Length, replacement);
        }
    }

    private static void ReplaceRangeInTextNodes(
        IReadOnlyList<Text> textNodes,
        int[] starts,
        int[] lengths,
        int matchStart,
        int matchLength,
        string replacement)
    {
        var matchEnd = matchStart + matchLength;
        var firstNode = -1;
        var lastNode = -1;
        for (var i = 0; i < textNodes.Count; i++)
        {
            var start = starts[i];
            var end = start + lengths[i];
            if (lengths[i] == 0)
            {
                continue;
            }

            if (firstNode < 0 && matchStart >= start && matchStart < end)
            {
                firstNode = i;
            }

            if (matchEnd > start && matchEnd <= end)
            {
                lastNode = i;
                break;
            }
        }

        if (firstNode < 0 || lastNode < 0)
        {
            return;
        }

        var firstText = textNodes[firstNode].Text ?? string.Empty;
        var lastText = textNodes[lastNode].Text ?? string.Empty;
        var firstOffset = Math.Clamp(matchStart - starts[firstNode], 0, firstText.Length);
        var lastOffset = Math.Clamp(matchEnd - starts[lastNode], 0, lastText.Length);
        var prefix = firstText[..firstOffset];
        var suffix = lastText[lastOffset..];

        textNodes[firstNode].Text = prefix + replacement + suffix;
        for (var i = firstNode + 1; i <= lastNode; i++)
        {
            textNodes[i].Text = string.Empty;
        }
    }

    private static void AppendDocumentWithPageBreak(string targetPath, string sourcePath)
    {
        using var target = WordprocessingDocument.Open(targetPath, true);
        using var source = WordprocessingDocument.Open(sourcePath, false);

        var targetBody = target.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Không đọc được nội dung file Word đích.");
        var sourceBody = source.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Không đọc được nội dung file Word nguồn.");

        InsertBeforeSectionProperties(targetBody, new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        foreach (var element in sourceBody.Elements().Where(e => e is not SectionProperties))
        {
            InsertBeforeSectionProperties(targetBody, element.CloneNode(true));
        }

        target.MainDocumentPart!.Document.Save();
    }

    private static void InsertBeforeSectionProperties(Body body, OpenXmlElement node)
    {
        var section = body.Elements<SectionProperties>().LastOrDefault();
        if (section is null)
        {
            body.Append(node);
            return;
        }

        body.InsertBefore(node, section);
    }
}
