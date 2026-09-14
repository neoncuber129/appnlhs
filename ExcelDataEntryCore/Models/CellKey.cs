namespace ExcelDataEntryApp.Models;

public readonly record struct CellKey(string SheetName, int RowIndex, int ColumnIndex);
