# Gufel.ExcelBuilder.Model

This project provides a flexible and extensible framework for building and exporting Excel files in .NET, with a focus on customizable column and value handling. It is designed to be used as the model layer for the `Gufel.ExcelBuilder` library.

## Key Components

### 1. `ExcelBuilder` Base Class

The `ExcelBuilder` class (located in the main project, not this model subfolder) is the core class responsible for creating and managing Excel files. It provides methods to:

- Add sheets from generic data collections.
- Customize column and value providers.
- Apply styles and formatting.
- Handle events for worksheet and column creation, and custom rendering.

#### Example Usage

```csharp
var builder = new ExcelBuilder();
builder.AddSheet("Sheet1", myDataList);
byte[] fileBytes = builder.BuildFile();
```

### 2. Interfaces

#### `IColumnProvider`

Located in `Model/Base/IColumnProvider.cs`, this interface abstracts how column metadata is provided to the builder.

```csharp
public interface IColumnProvider
{
    void SetSampleData(object? sampleData);
    List<ExcelColumnAttribute> GetColumns(Type dataType);
}
```

- **SetSampleData**: Allows the provider to inspect a sample object for dynamic column inference.
- **GetColumns**: Returns a list of `ExcelColumnAttribute` objects describing the columns for a given data type.

#### `IValueProvider`

Located in `Model/Base/IValueProvider.cs`, this interface abstracts how values are extracted from data objects for each column.

```csharp
public interface IValueProvider
{
    object? GetValue(ExcelColumnAttribute excelColumn, object classObject);
    Dictionary<string, object?> GetValues(List<ExcelColumnAttribute> excelColumns, object classObject);
}
```

- **GetValue**: Extracts a single value for a column from a data object.
- **GetValues**: Extracts all column values from a data object as a dictionary.

### 3. Attributes and Formats

#### `ExcelColumnAttribute`

Located in `Model/ExcelColumnAttribute.cs`, this attribute is used to annotate properties or fields in your data models to control how they are exported to Excel.

Key properties:
- `Name`: The column name in Excel.
- `HasValue`: Whether the column should be populated from the data source.
- `AsPersianDate`: Whether to convert dates to Persian format.
- `PersianDateFormat`: Custom format for Persian dates.
- `ColumnFormat`: Excel number format string.
- `SourceName`: The source property/field name in the data object.
- `Priority`: Order of the column.
- `SourceIsField`: Whether the source is a field (vs. property).

#### `ExcelColumnFormat`

Located in `Model/ExcelColumnFormat.cs`, this static class provides common Excel number format strings for convenience, such as:

- `OneDecimalPlace`
- `TwoDecimalPlace`
- `Percent`
- `ThousandSeparator`
- `NegativeRed`

## Extensibility

You can implement your own `IColumnProvider` or `IValueProvider` to customize how columns and values are determined and rendered in the Excel output.
