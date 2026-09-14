namespace ExcelDataEntryApp.Models;

public readonly record struct DropdownDependencyInfo(bool IsDependent, IReadOnlyList<int> ParentColumns)
{
    public static DropdownDependencyInfo Static { get; } = new(false, []);
}
