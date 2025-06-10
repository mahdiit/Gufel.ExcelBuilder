namespace Gufel.ExcelBuilder.Model
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Column name in excel
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Has value in datasource
        /// </summary>
        public bool HasValue { get; set; } = true;

        /// <summary>
        /// Convert date to persian date on render
        /// </summary>
        public bool AsPersianDate { get; set; }

        /// <summary>
        /// Date format apply on render
        /// </summary>
        public string? DateFormat { get; set; }

        /// <summary>
        /// Format apply to column in excel
        /// </summary>
        public string? ColumnFormat { get; set; }

        /// <summary>
        /// Source column
        /// </summary>
        public string? SourceName { get; set; }

        /// <summary>
        /// Priority
        /// </summary>
        public int Priority { get; set; }

        public bool SourceIsField { get; set; }
    }
}
