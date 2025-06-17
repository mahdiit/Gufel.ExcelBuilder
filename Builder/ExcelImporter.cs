using System.Reflection;
using Gufel.ExcelBuilder.ColumnProvider;
using Gufel.ExcelBuilder.Model.Base;
using OfficeOpenXml;

namespace Gufel.ExcelBuilder
{
    public sealed class ExcelImporter : IDisposable
    {
        private ExcelWorksheet? _currentWorksheet;
        private ExcelPackage _package;
        private ExcelWorkbook _workbook;
        private IColumnProvider _columnProvider = DefaultColumnProvider.Create();

        public ExcelImporter(Stream fileStream)
        {
            _package = new ExcelPackage(fileStream);
            _workbook = _package.Workbook;
            if (_workbook == null)
                throw new ExcelImportException("No workbook found on stream", "no.workbook.found");
        }

        #region Builder
        public static ExcelImporter FromStream(Stream excelStream)
        {
            return new ExcelImporter(excelStream);
        }
        public static ExcelImporter FromFile(string filePath)
        {
            return new ExcelImporter(File.OpenRead(filePath));
        }
        public static ExcelImporter FromBuffer(byte[] excelBuffer)
        {
            return new ExcelImporter(new MemoryStream(excelBuffer));
        }
        #endregion         

        public ExcelImporter SetColumnProvider(IColumnProvider columnProvider)
        {
            _columnProvider = columnProvider;
            return this;
        }

        private static Type? GetPropertyType(PropertyInfo? propertyInfo)
        {
            if (propertyInfo == null) return null;

            Type? propertyType = null;
            if (Nullable.GetUnderlyingType(propertyInfo.PropertyType) == null &&
                propertyInfo.PropertyType.GenericTypeArguments.Length > 0)
            {
                var gentType = propertyInfo.PropertyType.GenericTypeArguments[0].FullName;
                if (gentType != null)
                    propertyType = Type.GetType(gentType);
            }
            else
            {
                propertyType = propertyInfo.PropertyType;
            }

            return propertyType;

        }

        private static Type? GetFieldType(FieldInfo? fieldInfo)
        {
            if (fieldInfo == null) return null;

            Type? propertyType = null;
            if (Nullable.GetUnderlyingType(fieldInfo.FieldType) == null &&
                fieldInfo.FieldType.GenericTypeArguments.Length > 0)
            {
                var gentType = fieldInfo.FieldType.GenericTypeArguments[0].FullName;
                if (gentType != null)
                    propertyType = Type.GetType(gentType);
            }
            else
            {
                propertyType = fieldInfo.FieldType;
            }

            return propertyType;

        }

        public List<T> GetList<T>(string sheetName) where T : new()
        {
            if (_workbook.Worksheets.All(c => c.Name != sheetName))
                throw new ExcelImportException("Worksheet with this name not found", "worksheet.not.found");

            _currentWorksheet = _workbook.Worksheets[sheetName];
            
            var itemType = typeof(T);
            var columns = _columnProvider.GetColumns(itemType);

            var result = new List<T>();
            for (var rowNum = 2; rowNum <= _currentWorksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = _currentWorksheet.Cells[rowNum, 1, rowNum, _currentWorksheet.Dimension.End.Column];
                var item = Activator.CreateInstance<T>();
                foreach (var cell in wsRow)
                {
                    var column = columns[cell.Start.Column - 1];
                    Type? colType = null;

                    if (column.SourceIsField)
                    {
                        var field = itemType.GetField(column.SourceName!);
                        colType = GetFieldType(field);

                        if (field == null || colType == null)
                            continue;

                        field.SetValue(item, string.IsNullOrEmpty(cell.Text) ? null : Convert.ChangeType(cell.Text, colType));
                    }
                    else
                    {
                        var property = itemType.GetProperty(column.SourceName!);
                        colType = GetPropertyType(property);

                        if (property == null || colType == null)
                            continue;

                        property.SetValue(item, string.IsNullOrEmpty(cell.Text) ? null : Convert.ChangeType(cell.Text, colType));
                    }
                }
                result.Add(item);
            }
            return result;
        }

        public void Dispose()
        {
            _package.Dispose();
        }
    }
}
