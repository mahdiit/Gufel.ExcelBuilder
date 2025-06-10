using System.ComponentModel.DataAnnotations;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;

namespace Gufel.ExcelBuilder.ColumnProvider
{
    public class DefaultColumnProvider(bool onlyWithAttribute = false)
        : IColumnProvider
    {
        private Type? _dataType;

        public List<ExcelColumnAttribute> GetColumns(Type dataType)
        {
            _dataType = dataType;
            return FieldData().Concat(PropertyData()).ToList();
        }

        private List<ExcelColumnAttribute> PropertyData()
        {
            var props = _dataType!.GetProperties();
            var metadataType = _dataType.GetCustomAttributes(typeof(MetadataTypeAttribute), true)
                .OfType<MetadataTypeAttribute>().FirstOrDefault();

            var result = new List<ExcelColumnAttribute>();
            var metaDataProps = metadataType?.MetadataClassType.GetProperties();
            foreach (var prop in props)
            {
                if (onlyWithAttribute)
                {
                    var attr = prop.GetCustomAttributes(true).FirstOrDefault(c => c is ExcelColumnAttribute);
                    if (attr == null && metaDataProps != null)
                    {
                        var metaProp = metaDataProps.FirstOrDefault(x => x.Name == prop.Name);
                        if (metaProp != null)
                        {
                            attr = metaProp.GetCustomAttributes(true).FirstOrDefault(c => c is ExcelColumnAttribute);
                        }
                    }

                    if (attr == null) continue;

                    var excelAttr = (ExcelColumnAttribute)attr;
                    excelAttr.SourceName ??= prop.Name;
                    excelAttr.Name ??= excelAttr.SourceName;
                    excelAttr.SourceIsField = false;

                    result.Add(excelAttr);
                }
                else
                {
                    result.Add(new ExcelColumnAttribute { Name = prop.Name, SourceName = prop.Name, SourceIsField = false });
                }
            }

            return result.OrderBy(x => x.Priority).ToList();
        }

        private List<ExcelColumnAttribute> FieldData()
        {
            var props = _dataType!.GetFields();

            var result = new List<ExcelColumnAttribute>();
            foreach (var prop in props)
            {
                if (onlyWithAttribute)
                {
                    var attr = prop.GetCustomAttributes(true).FirstOrDefault(c => c is ExcelColumnAttribute);
                    if (attr == null) continue;

                    var excelAttr = (ExcelColumnAttribute)attr;
                    excelAttr.SourceName ??= prop.Name;
                    excelAttr.SourceIsField = true;

                    result.Add(excelAttr);
                }
                else
                {
                    result.Add(new ExcelColumnAttribute { Name = prop.Name, SourceName = prop.Name, SourceIsField = true });
                }
            }
            return result;
        }
    }
}
