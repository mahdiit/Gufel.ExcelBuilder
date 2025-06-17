using Gufel.ExcelBuilder.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gufel.ExcelBuilder.UnitTest
{
    public class ExcelBuilderTest
    {
        [Fact]
        public void WhenData_IsValid_ExcelMustSuccess()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Mehdi Yousefi");

            using var excelBuilder = new ExcelBuilder();
            List<TestModel> data = [
                new TestModel { ColDate = new DateTime(2024,2,2) , ColInt = 100, ColString = "Sample Data" }
            ];

            excelBuilder.AddSheet("Sheet1", data);
            var fileByte = excelBuilder.BuildFile();
            File.WriteAllBytes("C:\\Intel\\Logs\\samplefile.xlsx", fileByte);
            var excel = new ExcelPackage(new MemoryStream(fileByte));

            Assert.Equal(1, excel.Workbook?.Worksheets.Count);
            Assert.Equal("Sheet1", excel.Workbook?.Worksheets[0].Name);
            Assert.Equal("Int", excel.Workbook?.Worksheets[0].Cells[1,1].Value);
        }

        [Fact]
        public void WhenData_IsValid_ExcelMustImportSuccess()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Mehdi Yousefi");

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
