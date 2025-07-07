namespace Gufel.ExcelBuilder.Model;

public interface IDataReaderAdapter
{
    bool Read();
    object GetValue(int i);
    string GetName(int i);
    Type? GetFieldType(int i);
    int FieldCount { get; }
    bool IsDbNull(int i);
}