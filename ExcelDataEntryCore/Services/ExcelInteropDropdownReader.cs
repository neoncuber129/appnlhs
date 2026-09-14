using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataEntryApp.Services;

public sealed class ExcelInteropDropdownReader
{
    public IReadOnlyList<string> TryGetDropdownOptions(string workbookPath, string sheetName, int rowIndex, int columnIndex)
    {
        if (!ExcelInstallationProbe.IsInstalled())
        {
            return [];
        }

        Excel.Application? app = null;
        Excel.Workbook? workbook = null;
        Excel.Worksheet? worksheet = null;
        Excel.Range? cell = null;
        Excel.Validation? validation = null;

        try
        {
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            workbook = app.Workbooks.Open(workbookPath, ReadOnly: true);
            worksheet = workbook.Worksheets[sheetName] as Excel.Worksheet;
            if (worksheet is null)
            {
                return [];
            }

            cell = worksheet.Cells[rowIndex, columnIndex] as Excel.Range;
            validation = cell?.Validation;
            if (validation is null || validation.Type != (int)Excel.XlDVType.xlValidateList)
            {
                return [];
            }

            var formula1 = validation.Formula1 as string ?? string.Empty;
            return ResolveFormulaToValues(app, workbook, worksheet, formula1);
        }
        catch (Exception)
        {
            ExcelInstallationProbe.MarkUnavailable();
            return [];
        }
        finally
        {
            try
            {
                ReleaseComObject(validation);
                ReleaseComObject(cell);
                if (workbook is not null)
                {
                    workbook.Close(false);
                }

                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
                if (app is not null)
                {
                    app.Quit();
                }

                ReleaseComObject(app);
            }
            catch
            {
                // Bỏ qua lỗi dọn COM khi Excel/Office không khả dụng.
            }
        }
    }

    public IReadOnlyList<string> TryEvaluateListFormula(string workbookPath, string sheetName, string formula)
    {
        if (!ExcelInstallationProbe.IsInstalled() || string.IsNullOrWhiteSpace(formula))
        {
            return [];
        }

        Excel.Application? app = null;
        Excel.Workbook? workbook = null;
        Excel.Worksheet? worksheet = null;

        try
        {
            app = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            workbook = app.Workbooks.Open(workbookPath, ReadOnly: true);
            worksheet = workbook.Worksheets[sheetName] as Excel.Worksheet;
            if (worksheet is null)
            {
                return [];
            }

            workbook.Activate();
            worksheet.Activate();
            return ResolveFormulaToValues(app, workbook, worksheet, formula);
        }
        catch (Exception)
        {
            ExcelInstallationProbe.MarkUnavailable();
            return [];
        }
        finally
        {
            try
            {
                if (workbook is not null)
                {
                    workbook.Close(false);
                }

                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
                if (app is not null)
                {
                    app.Quit();
                }

                ReleaseComObject(app);
            }
            catch
            {
                // Bỏ qua lỗi dọn COM khi Excel/Office không khả dụng.
            }
        }
    }

    private static IReadOnlyList<string> ResolveFormulaToValues(Excel.Application app, Excel.Workbook workbook, Excel.Worksheet worksheet, string formula)
    {
        var normalized = formula.Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        if (!normalized.StartsWith("=", StringComparison.Ordinal)
            && !ValidationFormulaAdjuster.LooksLikeDynamicFormula(normalized))
        {
            return SplitInlineValues(normalized);
        }

        workbook.Activate();
        worksheet.Activate();

        var evaluateExpression = normalized.StartsWith("=", StringComparison.Ordinal)
            ? normalized
            : $"={normalized}";

        object? evaluated = null;
        Excel.Range? evaluatedRange = null;
        try
        {
            evaluated = app.Evaluate(evaluateExpression);
            evaluatedRange = evaluated as Excel.Range;
            if (evaluatedRange is not null)
            {
                return ReadRangeValues(evaluatedRange);
            }

            if (evaluated is object[,] matrix)
            {
                return ReadMatrixValues(matrix);
            }

            if (evaluated is string scalarText && !string.IsNullOrWhiteSpace(scalarText))
            {
                var inlineValues = SplitInlineValues(scalarText);
                if (inlineValues.Count > 0)
                {
                    return inlineValues;
                }
            }
        }
        catch
        {
            // Fallbacks below.
        }
        finally
        {
            ReleaseComObject(evaluatedRange);
        }

        return TryResolveFormulaToValues(workbook, worksheet, normalized);
    }

    private static IReadOnlyList<string> TryResolveFormulaToValues(Excel.Workbook workbook, Excel.Worksheet worksheet, string formula)
    {
        var target = formula[1..].Trim();

        Excel.Name? workbookName = null;
        Excel.Range? workbookNameRange = null;
        try
        {
            workbookName = workbook.Names.Item(target);
            workbookNameRange = workbookName?.RefersToRange;
            if (workbookNameRange is not null)
            {
                return ReadRangeValues(workbookNameRange);
            }
        }
        catch
        {
        }
        finally
        {
            ReleaseComObject(workbookNameRange);
            ReleaseComObject(workbookName);
        }

        Excel.Name? worksheetName = null;
        Excel.Range? worksheetNameRange = null;
        try
        {
            worksheetName = worksheet.Names.Item(target);
            worksheetNameRange = worksheetName?.RefersToRange;
            if (worksheetNameRange is not null)
            {
                return ReadRangeValues(worksheetNameRange);
            }
        }
        catch
        {
        }
        finally
        {
            ReleaseComObject(worksheetNameRange);
            ReleaseComObject(worksheetName);
        }

        Excel.Range? directRange = null;
        try
        {
            directRange = worksheet.Range[target];
            return ReadRangeValues(directRange);
        }
        catch
        {
            return [];
        }
        finally
        {
            ReleaseComObject(directRange);
        }
    }

    private static IReadOnlyList<string> ReadRangeValues(Excel.Range range)
    {
        var result = new List<string>();
        foreach (Excel.Range cell in range.Cells)
        {
            try
            {
                var text = cell.Value2?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    result.Add(text);
                }
            }
            finally
            {
                ReleaseComObject(cell);
            }
        }

        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static IReadOnlyList<string> ReadMatrixValues(object[,] matrix)
    {
        var result = new List<string>();
        foreach (var item in matrix)
        {
            var text = item?.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text))
            {
                result.Add(text);
            }
        }

        return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static IReadOnlyList<string> SplitInlineValues(string text)
    {
        return text.Trim('"')
            .Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
