using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;

namespace Gufel.ExcelBuilder.ColumnProvider
{
    public class DefaultColumnProvider(bool onlyColumnWithAttribute = true)
        : IColumnProvider
    {
        private static readonly Lazy<DefaultColumnProvider> Default = new(() => new DefaultColumnProvider());
        public static DefaultColumnProvider Create()
        {
            return Default.Value;
        }

        public List<ExcelColumnAttribute> GetColumns(Type dataType, object? data)
        {
            if (dataType == typeof(ExpandoObject))
            {
                return data == null ?
                    throw new ArgumentNullException("data in dynamic object must set")
                    : DynamicData(data);
            }

            var metadataType = dataType.GetCustomAttributes(typeof(MetadataTypeAttribute), true)
                .OfType<MetadataTypeAttribute>().FirstOrDefault();

            return FieldData(dataType, metadataType, onlyColumnWithAttribute)
                .Concat(PropertyData(dataType, metadataType, onlyColumnWithAttribute))
                .OrderBy(x => x.Priority)
                .ToList();
        }

        public static List<ExcelColumnAttribute> PropertyData(Type dataType, MetadataTypeAttribute? metadataType = null, bool onlyColWithAtt = true)
        {
            var props = dataType.GetProperties();
            var metaDataProps = metadataType?.MetadataClassType.GetProperties();

            var result = new List<ExcelColumnAttribute>();
            foreach (var prop in props)
            {
                if (onlyColWithAtt)
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

        public static List<ExcelColumnAttribute> FieldData(Type dataType, MetadataTypeAttribute? metadataType = null, bool onlyColWithAtt = true)
        {
            var props = dataType.GetFields();
            var metaDataProps = metadataType?.MetadataClassType.GetFields();

            var result = new List<ExcelColumnAttribute>();
            foreach (var prop in props)
            {
                if (onlyColWithAtt)
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

        public static List<ExcelColumnAttribute> DynamicData(object data)
        {
            return ((ExpandoObject)data)
                .Select(x => new ExcelColumnAttribute()
                {
                    Name = x.Key,
                    SourceName = x.Key,
                    SourceIsField = false
                }).ToList();
        }

        public static List<SqlExcelColumn> SqlReaderData(IDataReaderAdapter reader, List<ExcelColumnAttribute> columns)
        {
            var cols = new List<SqlExcelColumn>();
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var name = reader.GetName(i);
                var type = reader.GetFieldType(i);
                if (type == null)
                    continue;

                var col = columns.FirstOrDefault(x => x.SourceName == name) ?? new ExcelColumnAttribute()
                {
                    SourceName = name,
                    Priority = i,
                    Name = name
                };

                cols.Add(new SqlExcelColumn(i, type, col));
            }
            return cols;
        }
    }
}
