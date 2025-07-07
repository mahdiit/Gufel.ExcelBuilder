namespace Gufel.ExcelBuilder.Model.Base;

public interface IColumnProvider
{
    List<ExcelColumnAttribute> GetColumns(Type dataType, object? data);
}