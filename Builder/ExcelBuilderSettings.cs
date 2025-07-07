namespace Gufel.ExcelBuilder
{
    public record ExcelBuilderSettings
    {
        public bool UseDefaultStyle { get; set; } = true;
        public ExcelBuilderStyle? HeaderStyle { get; set; }
        public ExcelBuilderStyle? CellStyle { get; set; }
        public bool IsRtl { get; set; } = true;
        public string RowNumberColumnName { get; set; } = "ردیف";
        public bool HasRowNumber { get; set; } = false;
        public bool AutoFitColumns { get; set; } = true;

        internal int CellPadding => HasRowNumber ? 2 : 1;
    }
}
