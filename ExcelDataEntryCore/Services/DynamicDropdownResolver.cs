using OfficeOpenXml;

namespace ExcelDataEntryApp.Services;

/// <summary>
/// Giải dropdown phụ thuộc theo thứ tự ưu tiên (gần với Excel nhất có thể khi không mở Excel):
/// 1. OFFSET/MATCH/COUNTIF (file import KSK và nhiều template tương tự)
/// 2. Excel Interop Evaluate (nếu máy có Excel — xử lý INDIRECT và công thức khác)
/// 3. EPPlus đọc cột catalog từ công thức động (fallback rộng, có thể nhiều giá trị hơn Excel)
/// </summary>
public sealed class DynamicDropdownResolver
{
    private readonly ExcelInteropDropdownReader _interopReader = new();

    public IReadOnlyList<string> Resolve(
        ExcelPackage package,
        ExcelWorksheet dataWorksheet,
        string workbookPath,
        string sheetName,
        int rowIndex,
        int columnIndex,
        ExtendedListValidationRule rule,
        Func<int, int, string?> getCellValue,
        Func<string, IReadOnlyList<string>> resolveDynamicFormulaEpplus)
    {
        if (OffsetMatchDropdownResolver.TryResolve(
                package,
                rowIndex,
                columnIndex,
                rule,
                getCellValue,
                out var offsetValues))
        {
            return offsetValues;
        }

        var cellRange = rule.FindRange(rowIndex, columnIndex);
        if (cellRange is null)
        {
            return [];
        }

        var adjustedFormula = ValidationFormulaAdjuster.AdjustForRow(
            rule.Formula,
            cellRange.Value.StartRow,
            rowIndex);

        var interopValues = _interopReader.TryEvaluateListFormula(
            workbookPath,
            sheetName,
            adjustedFormula);
        if (interopValues.Count > 0)
        {
            return interopValues;
        }

        var epplusValues = resolveDynamicFormulaEpplus(adjustedFormula);
        return epplusValues.Count > 0 ? epplusValues : [];
    }
}
