namespace Gufel.ExcelBuilder.Model.Base;

public interface IColumnProvider
{
    void SetSampleData(object? sampleData);
    List<ExcelColumnAttribute> GetColumns(Type dataType);
}