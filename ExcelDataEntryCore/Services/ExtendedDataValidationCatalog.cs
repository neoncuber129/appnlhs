using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelDataEntryApp.Services;

/// <summary>
/// Đọc x14:dataValidation (công thức list động) mà EPPlus không expose qua DataValidations.
/// </summary>
public sealed class ExtendedDataValidationCatalog
{
    private static readonly XNamespace X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static readonly XNamespace Xm = "http://schemas.microsoft.com/office/excel/2006/main";

    private readonly Dictionary<string, IReadOnlyList<ExtendedListValidationRule>> _rulesBySheet =
        new(StringComparer.Ordinal);

    public IReadOnlyList<ExtendedListValidationRule> GetRules(string sheetName) =>
        _rulesBySheet.TryGetValue(sheetName, out var rules) ? rules : [];

    public ExtendedListValidationRule? FindRule(string sheetName, int rowIndex, int columnIndex)
    {
        foreach (var rule in GetRules(sheetName))
        {
            if (rule.Contains(rowIndex, columnIndex))
            {
                return rule;
            }
        }

        return null;
    }

    public void LoadFromWorkbook(string workbookPath)
    {
        _rulesBySheet.Clear();
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            return;
        }

        using var archive = ZipFile.OpenRead(workbookPath);
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (workbookEntry is null || relsEntry is null)
        {
            return;
        }

        var relIdToTarget = ReadRelationshipTargets(relsEntry);
        var workbookDoc = XDocument.Load(workbookEntry.Open());
        XNamespace main = workbookDoc.Root?.Name.Namespace ?? "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var sheetElement in workbookDoc.Descendants(main + "sheet"))
        {
            var sheetName = sheetElement.Attribute("name")?.Value ?? string.Empty;
            var relId = sheetElement.Attribute(relNs + "id")?.Value;
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(relId))
            {
                continue;
            }

            if (!relIdToTarget.TryGetValue(relId, out var target) || string.IsNullOrWhiteSpace(target))
            {
                continue;
            }

            var entryPath = target.StartsWith("/") ? target[1..] : $"xl/{target.TrimStart('/')}";
            var sheetEntry = archive.GetEntry(entryPath);
            if (sheetEntry is null)
            {
                continue;
            }

            var rules = ParseSheetValidations(sheetEntry);
            if (rules.Count > 0)
            {
                _rulesBySheet[sheetName] = rules;
            }
        }
    }

    private static Dictionary<string, string> ReadRelationshipTargets(ZipArchiveEntry relsEntry)
    {
        var relIdToTarget = new Dictionary<string, string>(StringComparer.Ordinal);
        var relsDoc = XDocument.Load(relsEntry.Open());
        foreach (var rel in relsDoc.Descendants().Where(e => e.Name.LocalName == "Relationship"))
        {
            var id = rel.Attribute("Id")?.Value;
            var target = rel.Attribute("Target")?.Value;
            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
            {
                relIdToTarget[id] = target;
            }
        }

        return relIdToTarget;
    }

    private static IReadOnlyList<ExtendedListValidationRule> ParseSheetValidations(ZipArchiveEntry sheetEntry)
    {
        var rules = new List<ExtendedListValidationRule>();
        var sheetDoc = XDocument.Load(sheetEntry.Open());
        foreach (var validationElement in sheetDoc.Descendants().Where(e => e.Name == X14 + "dataValidation"))
        {
            var formula = validationElement.Element(X14 + "formula1")?.Element(Xm + "f")?.Value?.Trim()
                ?? string.Empty;
            var sqref = validationElement.Element(Xm + "sqref")?.Value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(sqref))
            {
                continue;
            }

            var ranges = ParseSqrefRanges(sqref);
            if (ranges.Count == 0)
            {
                continue;
            }

            rules.Add(new ExtendedListValidationRule
            {
                Formula = formula,
                Sqref = sqref,
                Ranges = ranges,
                IsDynamic = ValidationFormulaAdjuster.LooksLikeDynamicFormula(formula),
                ParentColumns = ValidationFormulaAdjuster.ExtractSameSheetParentColumns(formula)
            });
        }

        return rules;
    }

    private static List<ValidationCellRange> ParseSqrefRanges(string sqref)
    {
        var ranges = new List<ValidationCellRange>();
        foreach (var token in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!TryParseRangeToken(token, out var range))
            {
                continue;
            }

            ranges.Add(range);
        }

        return ranges;
    }

    private static bool TryParseRangeToken(string token, out ValidationCellRange range)
    {
        range = default;
        var cleaned = token.Replace("$", string.Empty, StringComparison.Ordinal);
        var parts = cleaned.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 1)
        {
            if (!ExcelCellAddress.TryParseCell(parts[0], out var row, out var col))
            {
                return false;
            }

            range = new ValidationCellRange(row, row, col, col);
            return true;
        }

        if (parts.Length != 2
            || !ExcelCellAddress.TryParseCell(parts[0], out var startRow, out var startCol)
            || !ExcelCellAddress.TryParseCell(parts[1], out var endRow, out var endCol))
        {
            return false;
        }

        range = new ValidationCellRange(
            Math.Min(startRow, endRow),
            Math.Max(startRow, endRow),
            Math.Min(startCol, endCol),
            Math.Max(startCol, endCol));
        return true;
    }
}

public sealed class ExtendedListValidationRule
{
    public required string Sqref { get; init; }

    public required string Formula { get; init; }

    public required IReadOnlyList<ValidationCellRange> Ranges { get; init; }

    public required IReadOnlyList<int> ParentColumns { get; init; }

    public bool IsDynamic { get; init; }

    public bool Contains(int rowIndex, int columnIndex) =>
        Ranges.Any(r => rowIndex >= r.StartRow
            && rowIndex <= r.EndRow
            && columnIndex >= r.StartCol
            && columnIndex <= r.EndCol);

    public ValidationCellRange? FindRange(int rowIndex, int columnIndex) =>
        Ranges.FirstOrDefault(r => rowIndex >= r.StartRow
            && rowIndex <= r.EndRow
            && columnIndex >= r.StartCol
            && columnIndex <= r.EndCol);
}

public readonly record struct ValidationCellRange(int StartRow, int EndRow, int StartCol, int EndCol);
