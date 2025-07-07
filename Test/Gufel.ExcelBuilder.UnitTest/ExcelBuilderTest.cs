using Gufel.ExcelBuilder.Model;
using Moq;
using OfficeOpenXml;

namespace Gufel.ExcelBuilder.UnitTest
{
    public class ExcelBuilderTest
    {
        public ExcelBuilderTest()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Mehdi Yousefi");
        }

        [Fact]
        public void WhenData_IsValid_ExcelMustSuccess()
        {
            using var excelBuilder = new ExcelBuilder();
            List<TestModel> data = [
                new TestModel { ColDate = new DateTime(2024,2,2) , ColInt = 100, ColString = "Sample Data" }
            ];

            excelBuilder.AddSheet("Sheet1", data);
            var fileByte = excelBuilder.BuildFile();
            var excel = new ExcelPackage(new MemoryStream(fileByte));

            Assert.Equal(1, excel.Workbook?.Worksheets.Count);
            Assert.Equal("Sheet1", excel.Workbook?.Worksheets[0].Name);
            Assert.Equal("Int", excel.Workbook?.Worksheets[0].Cells[1, 1].Value);
        }

        [Fact]
        public void WhenData_IsValid_ExcelMustImportSuccess()
        {
            using var excelBuilder = new ExcelBuilder();
            List<TestModel> data = [
                new TestModel { ColDate = new DateTime(2024,2,2) , ColInt = 100, ColString = "Sample Data" }
            ];

            excelBuilder.AddSheet("Sheet1", data);
            var fileByte = excelBuilder.BuildFile();

            var excelImport = ExcelImporter.FromBuffer(fileByte);
            var result = excelImport.GetList<TestModel>("Sheet1");

            Assert.Equal(data.Count, result.Count);
            Assert.Equal(data[0].ColString, result[0].ColString);
        }

        [Fact]
        public void WhenSqlReader_GetData_ExcelMustSuccess()
        {
            var mock = new Mock<IDataReaderAdapter>();

            mock.SetupSequence(m => m.Read())
                .Returns(true)
                .Returns(false);

            mock.Setup(m => m.FieldCount).Returns(2);
            mock.Setup(m => m.GetName(0)).Returns("Id");
            mock.Setup(m => m.GetName(1)).Returns("Name");
            mock.Setup(m => m.GetFieldType(0)).Returns(typeof(int));
            mock.Setup(m => m.GetFieldType(1)).Returns(typeof(string));
            mock.Setup(m => m.IsDbNull(It.IsAny<int>())).Returns(false);
            mock.Setup(m => m.GetValue(0)).Returns(1);
            mock.Setup(m => m.GetValue(1)).Returns("Test");

            var excelBuilder = new ExcelBuilder();
            var fileByte = excelBuilder.AddSheet("SqlData", mock.Object).BuildFile();

            var excel = new ExcelPackage(new MemoryStream(fileByte));

            Assert.Equal(1, excel.Workbook?.Worksheets.Count);
            Assert.Equal("SqlData", excel.Workbook?.Worksheets[0].Name);
            Assert.Equal("Id", excel.Workbook?.Worksheets[0].Cells[1, 1].Value);
        }

        private record TestModel
        {
            [ExcelColumn(Name = "Int", Priority = 1)]
            public int ColInt { get; set; }

            [ExcelColumn(Name = "String", Priority = 3)]
            public string ColString { get; set; }

            [ExcelColumn(Name = "Date", Priority = 2, ColumnFormat = "yyyy/MM/dd")]
            public DateTime ColDate { get; set; }
        }
    }
}
