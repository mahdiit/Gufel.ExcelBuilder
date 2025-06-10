using Gufel.ExcelBuilder.Model;
using Gufel.ExcelBuilder.ValueProvider;
using System.Dynamic;

namespace Gufel.ExcelBuilder.UnitTest
{
    public class DefaultValueProviderTest
    {
        [Theory]
        [InlineData("IntField", true, 10)]
        [InlineData("IntProperty", false, 20)]
        public void ValueOfModel_WhenAttributeIsValid_MustEqualToOriginalValue(string name, bool isField, int originalValue)
        {
            var model = new TestModel() { IntField = 10, IntProperty = 20 };
            var provider = new DefaultValueProvider();
            var attr = new ExcelColumnAttribute() { SourceName = name, SourceIsField = isField, Name = name };

            var value = provider.GetValue(attr, model);
            Assert.Equal(value, originalValue);
        }

        [Fact]
        public void ValuesOfModel_WhenAttributeIsValid_MustEqualToOriginalValues()
        {
            var model = new TestModel() { IntField = 10, IntProperty = 20 };
            var provider = new DefaultValueProvider();
            var attrField = new ExcelColumnAttribute() { SourceName = "IntField", SourceIsField = true, Name = "IntField" };
            var attrProperty = new ExcelColumnAttribute() { SourceName = "IntProperty", SourceIsField = false, Name = "IntProperty" };

            var values = provider.GetValues([attrField, attrProperty], model);

            Assert.Equal(values["IntField"], 10);
            Assert.Equal(values["IntProperty"], 20);
        }

        [Fact]
        public void ValuesOfExpando_WhenAttributeIsValid_MustEqualToOriginalValues()
        {
            dynamic model = new ExpandoObject();
            model.IntField = 10;
            model.IntProperty = 20;

            var provider = new DefaultValueProvider();
            var attrField = new ExcelColumnAttribute() { SourceName = "IntField", SourceIsField = false, Name = "IntField" };
            var attrProperty = new ExcelColumnAttribute() { SourceName = "IntProperty", SourceIsField = false, Name = "IntProperty" };

            var values = provider.GetValues(new List<ExcelColumnAttribute>() { attrProperty, attrField }, model);

            Assert.Equal(values["IntField"], 10);
            Assert.Equal(values["IntProperty"], 20);
        }

        private class TestModel
        {
            public int IntProperty { get; set; }
            public int IntField;
        }
    }
}
