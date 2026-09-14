namespace ExcelDataEntryApp.Models;

public sealed class SheetImportConfig
{
    public required string SheetName { get; init; }

    public bool IsHidden { get; init; }

    public int HeaderFirstRow { get; init; } = 1;

    public int HeaderLastRow { get; init; } = 3;

    public int FirstDataRow { get; init; } = 4;

    public int SampleRow { get; init; }

    public bool IsNameSheet { get; init; }

    public int NameColumnIndex { get; init; } = 2;
}
