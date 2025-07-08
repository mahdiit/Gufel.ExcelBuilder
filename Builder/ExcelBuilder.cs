using System.Collections;
using System.Drawing;
using Gufel.Date;
using Gufel.ExcelBuilder.ColumnProvider;
using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.Model.Base;
using Gufel.ExcelBuilder.ValueProvider;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Gufel.ExcelBuilder
{
    public delegate bool RenderColumn(string column, object? value, ExcelRange excelColumn, Dictionary<string, object?> rowData);
    public delegate void CreateWorksheet(ExcelWorksheet ws);
    public delegate void CreateColumn(ExcelColumn column);

    public sealed class ExcelBuilder : IDisposable
    {
        public ExcelBuilder()
        {
            Settings = new ExcelBuilderSettings();
            _memoryStream = new MemoryStream();
            _xlsx = new ExcelPackage(_memoryStream);
        }

        private readonly MemoryStream _memoryStream;
        private readonly ExcelPackage _xlsx;

        public event CreateWorksheet? OnCreateWorksheet;
        public event CreateColumn? OnCreateColumn;
        public event RenderColumn? OnRenderColumn;

        private IColumnProvider _columnProvider = DefaultColumnProvider.Create();
        private IValueProvider _valueProvider = DefaultValueProvider.Create();

        public ExcelBuilderSettings Settings { get; private set; }

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

        public ExcelBuilder SetSettings(ExcelBuilderSettings settings)
        {
            Settings = settings;
            return this;
        }

        public ExcelBuilderStyle CreateStyle(string name)
        {
            return new ExcelBuilderStyle(name, _xlsx.Workbook.Styles.CreateNamedStyle(name));
        }

        public ExcelBuilder AddSheet<T>(string name, List<T> data)
        {
            var colInfoList = _columnProvider.GetColumns(typeof(T), data.FirstOrDefault());
            return AddSheet(name, colInfoList, data);

        }

        public ExcelBuilder AddSheet(string name, IEnumerable data)
        {
            var enumerable = data as object[] ?? data.Cast<object>().ToArray();
            var enumerator = enumerable.GetEnumerator();
            if (!enumerator.MoveNext())
            {
                ((IDisposable)enumerator).Dispose();
                return this;
            }

            var first = enumerator.Current ?? throw new ExcelBuildException("Empty data", "empty.data");

            var fType = first.GetType();
            var colInfoList = _columnProvider.GetColumns(fType, first);
            return AddSheet(name, colInfoList, enumerable);
        }

        public ExcelBuilder AddSheet(string name, IDataReaderAdapter reader)
        {
            var columns = _columnProvider.GetColumns(typeof(IDataReaderAdapter), null);
            var sqlColumns = DefaultColumnProvider.SqlReaderData(reader, columns ?? []);
            var ws = PrepareSheet(name, sqlColumns.Select(x => x.Column).OrderBy(x => x.Priority).ToList());

            var totalCount = 0;
            while (reader.Read())
            {
                totalCount++;
                if (Settings.HasRowNumber)
                {
                    ws.Cells[totalCount + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[totalCount + 1, 1].Value = totalCount;
                }

                foreach (var sqlColumn in sqlColumns)
                {
                    var col = ws.Cells[totalCount + 1, sqlColumn.Column.Priority + Settings.CellPadding];
                    var sqlValue = reader.GetValue(sqlColumn.SqlPriority);
                    var objValue = (sqlValue != DBNull.Value) ? Convert.ChangeType(sqlValue, sqlColumn.SqlType) : null;
                    var isRenderFinish = false;
                    if (OnRenderColumn != null)
                    {
                        var objVal = sqlColumn.Column.HasValue ? objValue : null;
                        isRenderFinish = OnRenderColumn(sqlColumn.Column.SourceName!, objVal, col, null);
                    }

                    if (isRenderFinish || !sqlColumn.Column.HasValue) continue;

                    col.Value = objValue;
                    var columnFormat = sqlColumn.Column.ColumnFormat;

                    if (sqlColumn.Column.AsPersianDate && objValue is DateTime dateTime)
                    {
                        var dateFormat = sqlColumn.Column.PersianDateFormat;
                        var vDate = new VDate(dateTime);
                        col.Value = vDate.ToString(dateFormat ?? "$yyyy/$MM/$dd");
                    }
                    else if (objValue is DateTime)
                    {
                        col.Style.Numberformat.Format = columnFormat ?? "yyyy/MM/dd HH:mm:ss";
                    }
                    else if (columnFormat != null)
                        col.Style.Numberformat.Format = columnFormat;
                }

                if (totalCount == 1_000_000)
                {
                    break;
                }
            }

            DoneSheet(ws);

            return this;
        }

        private ExcelBuilder AddSheet(string name, List<ExcelColumnAttribute> colInfoList, IEnumerable data)
        {
            var ws = PrepareSheet(name, colInfoList);

            var totalCount = 0;

            foreach (var row in data)
            {
                totalCount++;

                if (row == null)
                    continue;

                if (Settings.HasRowNumber)
                {
                    ws.Cells[totalCount + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[totalCount + 1, 1].Value = totalCount;
                }

                var colsData = _valueProvider.GetValues(colInfoList, row);
                var colsIndex = 0;
                foreach (var colData in colsData)
                {
                    var col = ws.Cells[totalCount + 1, colsIndex + Settings.CellPadding];

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

                        if (colInfoList[colsIndex].AsPersianDate && colData.Value is DateTime dateTime)
                        {
                            var dateFormat = colInfoList[colsIndex].PersianDateFormat;
                            var vDate = new VDate(dateTime);
                            col.Value = vDate.ToString(dateFormat ?? "$yyyy/$MM/$dd");
                        }
                        else if (colData.Value is DateTime)
                        {
                            col.Style.Numberformat.Format = columnFormat ?? "yyyy/MM/dd HH:mm:ss";
                        }
                        else if (columnFormat != null)
                            col.Style.Numberformat.Format = columnFormat;
                    }

                    colsIndex++;
                }

                if (totalCount == 1_000_000)
                {
                    break;
                }
            }

            DoneSheet(ws);

            return this;
        }

        private void DoneSheet(ExcelWorksheet ws)
        {
            if (Settings.AutoFitColumns)
                ws.Cells.AutoFitColumns();

            ws.View.PageLayoutView = false;
            ws.View.RightToLeft = Settings.IsRtl;
        }

        private ExcelWorksheet PrepareSheet(string sheetName, List<ExcelColumnAttribute> colInfoList)
        {
            var ws = _xlsx.Workbook.Worksheets.Add(sheetName);

            OnCreateWorksheet?.Invoke(ws);

            CheckStyles();

            if (Settings.CellStyle != null)
                ws.Cells.StyleName = Settings.CellStyle.Name;

            if (Settings.HeaderStyle != null)
                ws.Cells[1, 1].StyleName = Settings.HeaderStyle.Name;

            if (Settings.HasRowNumber)
                ws.Cells[1, 1].Value = Settings.RowNumberColumnName;

            for (var i = 0; i < colInfoList.Count; i++)
            {
                if (Settings.HeaderStyle != null)
                    ws.Cells[1, i + Settings.CellPadding].StyleName = Settings.HeaderStyle.Name;

                ws.Cells[1, i + Settings.CellPadding].Value = colInfoList[i].Name;

                OnCreateColumn?.Invoke(ws.Column(i + 2));
            }

            return ws;
        }

        private void CheckStyles()
        {
            if (Settings is { UseDefaultStyle: true, HeaderStyle: null })
            {
                Settings.HeaderStyle = CreateStyle("CmCellStyle");
                Settings.HeaderStyle.Object.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Settings.HeaderStyle.Object.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Settings.HeaderStyle.Object.Style.WrapText = false;
                Settings.HeaderStyle.Object.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Settings.HeaderStyle.Object.Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);
            }

            if (Settings is not { UseDefaultStyle: true, CellStyle: null }) return;

            Settings.CellStyle = CreateStyle("CmHeaderStyle");
            Settings.CellStyle.Object.Style.Font.Name = "Tahoma";
            Settings.CellStyle.Object.Style.Font.Size = 9.5f;
        }

        public byte[] BuildFile()
        {
            _xlsx.Save();
            return _memoryStream.ToArray();
        }

        public void Dispose()
        {
            _xlsx.Dispose();
            _memoryStream.Dispose();
        }
    }
}
