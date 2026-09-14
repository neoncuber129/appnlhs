using System.IO;
using ExcelDataEntryApp.Models;
using OfficeOpenXml;

namespace ExcelDataEntryApp.Services;

public static class ImportFileExportService
{
    public static ImportFileExportResult Export(
        string sourceWorkbookPath,
        string outputWorkbookPath,
        IReadOnlyList<ImportFileSheetPlan> sheetPlans)
    {
        if (string.IsNullOrWhiteSpace(sourceWorkbookPath) || !File.Exists(sourceWorkbookPath))
        {
            throw new FileNotFoundException("Không tìm thấy file Excel nguồn.", sourceWorkbookPath);
        }

        if (sheetPlans.Count == 0)
        {
            throw new InvalidOperationException("Không có sheet nào để xuất.");
        }

        var outputDirectory = Path.GetDirectoryName(outputWorkbookPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.Copy(sourceWorkbookPath, outputWorkbookPath, overwrite: true);

        ExcelPackage.License.SetNonCommercialPersonal("ExcelDataEntryApp");
        using var package = new ExcelPackage(new FileInfo(outputWorkbookPath));

        var filledCells = 0;
        var processedRows = 0;

        foreach (var plan in sheetPlans)
        {
            if (plan.LastDataRow < plan.FirstDataRow || plan.ColumnIndexes.Count == 0 || plan.SampleRow < 1)
            {
                continue;
            }

            var worksheet = package.Workbook.Worksheets[plan.SheetName]
                ?? throw new InvalidOperationException($"Không tìm thấy sheet '{plan.SheetName}'.");

            var skipColumns = plan.SkipSampleColumnIndexes.ToHashSet();
            var sampleValues = plan.ColumnIndexes
                .Where(column => !skipColumns.Contains(column))
                .ToDictionary(
                    column => column,
                    column => worksheet.Cells[plan.SampleRow, column].Text?.Trim() ?? string.Empty);

            for (var row = plan.FirstDataRow; row <= plan.LastDataRow; row++)
            {
                processedRows++;
                var useFemaleOnlyColumns = ShouldUseFemaleOnlyColumnsForRow(package, plan, row);

                foreach (var column in plan.ColumnIndexes)
                {
                    if (skipColumns.Contains(column) && !useFemaleOnlyColumns)
                    {
                        continue;
                    }

                    var current = worksheet.Cells[row, column].Text?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(current))
                    {
                        continue;
                    }

                    if (!sampleValues.TryGetValue(column, out var sampleValue) || string.IsNullOrWhiteSpace(sampleValue))
                    {
                        if (skipColumns.Contains(column) && useFemaleOnlyColumns)
                        {
                            sampleValue = worksheet.Cells[plan.SampleRow, column].Text?.Trim() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(sampleValue))
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }

                    worksheet.Cells[row, column].Value = sampleValue;
                    filledCells++;
                }
            }
        }

        package.Save();

        return new ImportFileExportResult
        {
            FilledCellCount = filledCells,
            ProcessedRowCount = processedRows,
            ProcessedSheetCount = sheetPlans.Count,
            OutputPath = outputWorkbookPath
        };
    }

    private static bool ShouldUseFemaleOnlyColumnsForRow(ExcelPackage package, ImportFileSheetPlan plan, int row)
    {
        if (plan.GenderColumnIndex <= 0)
        {
            return false;
        }

        var genderSheetName = string.IsNullOrWhiteSpace(plan.GenderSheetName)
            ? plan.SheetName
            : plan.GenderSheetName;
        var genderWorksheet = package.Workbook.Worksheets[genderSheetName];
        if (genderWorksheet is null)
        {
            return false;
        }

        var genderRow = plan.GenderSheetFirstDataRow + (row - plan.FirstDataRow);
        var genderText = genderWorksheet.Cells[genderRow, plan.GenderColumnIndex].Text?.Trim() ?? string.Empty;
        return HsskExportGenderHelper.ShouldUseFemaleOnlyColumns(genderText);
    }
}
