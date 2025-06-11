using OfficeOpenXml.Style.XmlAccess;

namespace Gufel.ExcelBuilder
{
    public record ExcelBuilderSettings
    {
        public bool UseDefaultStyle { get; set; } = true;
        public ExcelNamedStyleXml? HeaderStyle { get; set; }
        public ExcelNamedStyleXml? CellStyle { get; set; }
        public string CellStyleName { get; private set; } = "CmCellStyle";
        public string HeaderStyleName { get; private set; } = "CmHeaderStyle";
        public bool IsRtl { get; set; } = true;
        public string RowNumberColumnName { get; set; } = "ردیف";
        public bool HasRowNumber { get; set; } = false;
    }
}
