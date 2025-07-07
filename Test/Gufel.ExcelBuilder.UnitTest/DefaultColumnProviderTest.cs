using Gufel.ExcelBuilder.ColumnProvider;
using Gufel.ExcelBuilder.Model;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;

namespace Gufel.ExcelBuilder.UnitTest
{
    public class DefaultColumnProviderTest
    {
        [Fact]
        public void ProviderColumn_WhenTypeIsCorrect_MustReturnValidAttributes()
        {
            var provider = new DefaultColumnProvider(onlyColumnWithAttribute: true);

            var cols = provider.GetColumns(typeof(TestModel), null);

            Assert.NotNull(cols.Find(x => x is { Name: "IntProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "StringProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "BoolProperty", SourceIsField: true }));
        }

        [Fact]
        public void ProviderColumn_WhenTypeIsMetadata_MustReturnValidAttributes()
        {
            var provider = new DefaultColumnProvider(onlyColumnWithAttribute: true);

            var cols = provider.GetColumns(typeof(TestModelSimple), null);

            Assert.NotNull(cols.Find(x => x is { Name: "IntProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "StringProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "BoolProperty", SourceIsField: true }));
        }

        [Fact]
        public void ProviderColumn_WhenTypeIsDynamic_MustReturnValidAttributes()
        {
            var provider = new DefaultColumnProvider(onlyColumnWithAttribute: false);
            dynamic data = new ExpandoObject();
            data.IntProperty = 1;
            data.StringProperty = "";
            data.BoolProperty = false;

            List<ExcelColumnAttribute> cols = provider.GetColumns(typeof(ExpandoObject), data);

            Assert.NotNull(cols.Find(x => x is { Name: "IntProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "StringProperty", SourceIsField: false }));
            Assert.NotNull(cols.Find(x => x is { Name: "BoolProperty", SourceIsField: false }));
        }

        private record TestModel(int IntProperty, string StringProperty, bool BoolProperty)
        {
            [ExcelColumn]
            public int IntProperty { get; set; } = IntProperty;

            [ExcelColumn]
            public string StringProperty { get; set; } = StringProperty;

            [ExcelColumn] public bool BoolProperty = BoolProperty;
        }

        private class TestModelSimpleMetaData
        {
            [ExcelColumn]
            public int IntProperty { get; init; }

            [ExcelColumn]
            public string StringProperty { get; init; }

            [ExcelColumn] public bool BoolProperty;
        }

        [MetadataType(typeof(TestModelSimpleMetaData))]
        private class TestModelSimple(int intProperty, string stringProperty, bool boolProperty)
        {
            public int IntProperty { get; set; } = intProperty;

            public string StringProperty { get; set; } = stringProperty;

            public bool BoolProperty = boolProperty;
        }
    }
}
