using System.Dynamic;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;

namespace Gufel.ExcelBuilder.ValueProvider
{
    public class DefaultValueProvider : IValueProvider
    {
        public object? GetValue(ExcelColumnAttribute excelColumn, object classObject)
        {
            return excelColumn.SourceIsField
                ? classObject.GetType().GetField(excelColumn.SourceName!)?.GetValue(classObject)
                : classObject.GetType().GetProperty(excelColumn.SourceName!)?.GetValue(classObject, null);
        }

        public Dictionary<string, object?> GetValues(List<ExcelColumnAttribute> excelColumns, object classObject)
        {
            var result = new Dictionary<string, object?>();
            if (classObject is ExpandoObject)
            {
                var sm = (IDictionary<string, object>)classObject;
                foreach (var item in excelColumns.Select(x => x.SourceName))
                {
                    result.Add(item!, sm[item!]);
                }
                return result;
            }

            foreach (var name in excelColumns)
            {
                result.Add(name.SourceName!, GetValue(name, classObject));
            }
            return result;
        }
    }
}
