using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;

namespace Gufel.ExcelBuilder.ColumnProvider
{
    public class ManualColumnProvider(List<ExcelColumnAttribute> columns) : IColumnProvider
    {
        public List<ExcelColumnAttribute> GetColumns(Type dataType)
        {
            return columns;
        }

        public void SetSampleData(object? sampleData)
        {
            
        }
    }
}
