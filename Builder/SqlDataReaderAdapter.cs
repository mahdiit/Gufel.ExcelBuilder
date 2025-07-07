using Gufel.ExcelBuilder.Model;
using Microsoft.Data.SqlClient;

namespace Gufel.ExcelBuilder;

public class SqlDataReaderAdapter(SqlDataReader reader) : IDataReaderAdapter
{
    public bool Read() => reader.Read();
    public object GetValue(int i) => reader.GetValue(i);
    public string GetName(int i) => reader.GetName(i);
    public Type? GetFieldType(int i) => reader.GetFieldType(i);
    public int FieldCount => reader.FieldCount;
    public bool IsDbNull(int i) => reader.IsDBNull(i);
}