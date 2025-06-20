﻿using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;

namespace Gufel.ExcelBuilder.ColumnProvider
{
    public class DefaultColumnProvider(bool onlyWithAttribute = true)
        : IColumnProvider
    {
        object? _sampleData;
        private static readonly Lazy<DefaultColumnProvider> Default = new(() => new DefaultColumnProvider());
        public static DefaultColumnProvider Create()
        {
            return Default.Value;
        }

        private Type? _dataType;
        private MetadataTypeAttribute? _metadataType;

        public List<ExcelColumnAttribute> GetColumns(Type dataType)
        {
            if (dataType == typeof(ExpandoObject))
            {
                return _sampleData == null ? 
                    throw new ArgumentNullException("sample data in dynamic object must set") 
                    : DynamicData();
            }

            _dataType = dataType;
            _metadataType = _dataType.GetCustomAttributes(typeof(MetadataTypeAttribute), true)
                .OfType<MetadataTypeAttribute>().FirstOrDefault();

            return FieldData().Concat(PropertyData())
                .OrderBy(x => x.Priority)
                .ToList();
        }

        private List<ExcelColumnAttribute> PropertyData()
        {
            var props = _dataType!.GetProperties();
            var metaDataProps = _metadataType?.MetadataClassType.GetProperties();

            var result = new List<ExcelColumnAttribute>();
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
            var metaDataProps = _metadataType?.MetadataClassType.GetFields();

            var result = new List<ExcelColumnAttribute>();
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

        private List<ExcelColumnAttribute> DynamicData()
        {
            return ((ExpandoObject)_sampleData!)
                .Select(x => new ExcelColumnAttribute()
                {
                    Name = x.Key,
                    SourceName = x.Key,
                    SourceIsField = false
                }).ToList();
        }

        public void SetSampleData(object? sampleData)
        {
            _sampleData = sampleData;
        }
    }
}
