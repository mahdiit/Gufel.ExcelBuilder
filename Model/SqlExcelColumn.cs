namespace Gufel.ExcelBuilder.Model;

public class SqlExcelColumn(int sqlPriority, Type sqlType, ExcelColumnAttribute column)
{
    public int SqlPriority { get; } = sqlPriority;
    public Type SqlType { get; } = sqlType;
    public ExcelColumnAttribute Column { get; } = column;
}