using System.Collections;
using System.Drawing;
using System.Dynamic;
using Gufel.Date;
using Gufel.ExcelBuilder.ColumnProvider;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;
using Gufel.ExcelBuilder.ValueProvider;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace Gufel.ExcelBuilder
{
    public delegate bool RenderColumn(string column, object? value, ExcelRange excelColumn, Dictionary<string, object?> rowData);
    public delegate void CreateWorksheet(ExcelWorksheet ws);
    public delegate void CreateColumn(ExcelColumn column);

    public sealed class ExcelBuilder
    {
        private MemoryStream? _memoryStream;
        private ExcelPackage? _xlsx;

        public event CreateWorksheet? OnCreateWorksheet;
        public event CreateColumn? OnCreateColumn;
        public event RenderColumn? OnRenderColumn;

        public string? RowNumberColumnName { get; set; } = null;
        private bool HasRowNumber => !string.IsNullOrEmpty(RowNumberColumnName);

        private IColumnProvider _columnProvider = new DefaultColumnProvider();
        private IValueProvider _valueProvider = new DefaultValueProvider();

        public ExcelBuilder SetColumnProvider(IColumnProvider provider)
        {
            _columnProvider = provider;
            return this;
        }

        public ExcelBuilder SetValueProvider(IValueProvider provider)
        {
            _valueProvider = provider;
            return this;
        }

        public ExcelBuilder Create(string? rowNumber = null, bool setDefaultStyle = false)
        {
            RowNumberColumnName = rowNumber;

            if (_xlsx != null) return this;

            _memoryStream = new MemoryStream();
            _xlsx = new ExcelPackage(_memoryStream);

            if (!setDefaultStyle) return this;

            var headerStyle = CreateStyle("CmHeaderStyle");
            HeaderStyleName = "CmHeaderStyle";
            headerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            headerStyle.Style.WrapText = false;
            headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerStyle.Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);

            var cellStyle = CreateStyle("CmCellStyle");
            CellStyleName = "CmCellStyle";
            cellStyle.Style.Font.Name = "Tahoma";
            cellStyle.Style.Font.Size = 9.5f;
            return this;
        }

        public string? CellStyleName { get; set; }
        public string? HeaderStyleName { get; set; }
        public bool IsRtl { get; set; } = true;

        public ExcelNamedStyleXml CreateStyle(string name)
        {
            if (_xlsx == null)
                throw new ExcelBuildException("workbook is empty", "empty.workbook");

            return _xlsx.Workbook.Styles.CreateNamedStyle(name);
        }

        public ExcelBuilder AddSheet<T>(string name, List<T> data, bool autoFitColumns = true)
        {
            if (_xlsx == null)
                throw new ExcelBuildException("workbook is empty", "empty.workbook");

            var colInfoList = _columnProvider.GetColumns(typeof(T));
            return AddSheet(name, colInfoList, data, autoFitColumns);

        }

        public ExcelBuilder AddSheet(string name, IEnumerable data, bool autoFitColumns = true)
        {
            if (_xlsx == null)
                throw new ExcelBuildException("workbook is empty", "empty.workbook");

            var enumerable = data as object[] ?? data.Cast<object>().ToArray();
            var enumerator = enumerable.GetEnumerator();
            if (!enumerator.MoveNext())
            {
                ((IDisposable)enumerator).Dispose();
                return this;
            }

            var first = enumerator.Current;
            if (first == null)
                throw new ExcelBuildException("Empty data", "empty.data");

            var fType = first.GetType();
            var colInfoList = _columnProvider.GetColumns(fType);
            if (fType != typeof(ExpandoObject)) return AddSheet(name, colInfoList, enumerable, autoFitColumns);

            var expandoDict = (IEnumerable<ExpandoObject>)enumerable;
            var colInfoList2 = expandoDict.First()
                .Select(x => new ExcelColumnAttribute()
                {
                    Name = x.Key,
                    SourceName = x.Key,
                    SourceIsField = false
                }).ToList();

            colInfoList = (from all in colInfoList2
                           join cols in colInfoList on all.SourceName equals cols.SourceName
                           select cols).ToList();

            return AddSheet(name, colInfoList, enumerable, autoFitColumns);
        }


        private ExcelBuilder AddSheet(string name, List<ExcelColumnAttribute> colInfoList, IEnumerable data, bool autoFitColumns = true)
        {
            if (_xlsx == null) return this;

            var ws = _xlsx.Workbook.Worksheets.Add(name);

            OnCreateWorksheet?.Invoke(ws);

            if (!string.IsNullOrEmpty(CellStyleName))
                ws.Cells.StyleName = CellStyleName;

            if (!string.IsNullOrEmpty(HeaderStyleName) && HasRowNumber)
                ws.Cells[1, 1].StyleName = HeaderStyleName;

            if (HasRowNumber)
                ws.Cells[1, 1].Value = RowNumberColumnName;

            var totalCount = 0;
            var cellPadding = HasRowNumber ? 2 : 1;

            for (var i = 0; i < colInfoList.Count; i++)
            {
                if (!string.IsNullOrEmpty(HeaderStyleName))
                    ws.Cells[1, i + cellPadding].StyleName = HeaderStyleName;

                ws.Cells[1, i + cellPadding].Value = colInfoList[i].Name;

                OnCreateColumn?.Invoke(ws.Column(i + 2));
            }

            foreach (var row in data)
            {
                totalCount++;

                if (row == null)
                    continue;

                if (HasRowNumber)
                {
                    ws.Cells[totalCount + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[totalCount + 1, 1].Value = totalCount;
                }

                var colsData = _valueProvider.GetValues(colInfoList, row);
                var colsIndex = 0;
                foreach (var colData in colsData)
                {
                    var col = ws.Cells[totalCount + 1, colsIndex + cellPadding];

                    var isRenderFinish = false;
                    if (OnRenderColumn != null)
                    {
                        var objVal = colInfoList[colsIndex].HasValue ? colData.Value : null;
                        isRenderFinish = OnRenderColumn(colData.Key, objVal, col, colsData);
                    }

                    if (!isRenderFinish && colInfoList[colsIndex].HasValue)
                    {
                        col.Value = colData.Value;
                        var columnFormat = colInfoList[colsIndex].ColumnFormat;
                        var dateFormat = colInfoList[colsIndex].DateFormat;

                        if (colInfoList[colsIndex].AsPersianDate && colData.Value is DateTime dateTime)
                        {
                            var vDate = new VDate(dateTime);
                            col.Value = vDate.ToString(dateFormat ?? "$yyyy/$MM/$dd");
                        }
                        else if (columnFormat == null && colData.Value is DateTime)
                        {
                            col.Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                        }
                        else if (columnFormat != null)
                            col.Style.Numberformat.Format = colInfoList[colsIndex].ColumnFormat;
                    }

                    colsIndex++;
                }

                if (totalCount == 1000000)
                {
                    break;
                }
            }

            if (autoFitColumns)
                ws.Cells.AutoFitColumns();

            ws.View.PageLayoutView = false;
            ws.View.RightToLeft = IsRtl;

            return this;
        }

        public byte[] Build()
        {
            if (_xlsx == null || _memoryStream == null)
                throw new ExcelBuildException("document is empty", "empty.document");

            _xlsx.Save();
            return _memoryStream.ToArray();
        }

        public void Dispose()
        {
            _xlsx?.Dispose();
            _memoryStream?.Dispose();
        }
    }
}
