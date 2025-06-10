namespace Gufel.ExcelBuilder.Model.Base;

public interface IValueProvider
{
    object? GetValue(ExcelColumnAttribute excelColumn, object classObject);
    Dictionary<string, object?> GetValues(List<ExcelColumnAttribute> excelColumns, object classObject);
}