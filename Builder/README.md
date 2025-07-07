# Gufel.ExcelBuilder

A powerful and flexible Excel file builder library for .NET applications that makes it easy to create and customize Excel files programmatically.

## Features

- Create Excel files with multiple worksheets
- Custom column formatting and styling
- Support for Persian date formatting
- Customizable cell styles and headers
- Right-to-Left (RTL) support
- Auto-fit columns
- Row numbering
- Custom value providers
- Custom column providers
- Event-based customization
- Support for dynamic data
- Memory-efficient processing
- Thread-safe operations
- Support sql data reader

## Installation

```bash
dotnet add package Gufel.ExcelBuilder
```

## Usage

### Basic Example

```csharp
using Gufel.ExcelBuilder;

// Create a list of data
var data = new List<Person>
{
    new Person { Name = "John", Age = 30, BirthDate = new DateTime(1993, 1, 1) },
    new Person { Name = "Jane", Age = 25, BirthDate = new DateTime(1998, 5, 15) }
};

// Create and configure the Excel builder
using var builder = new ExcelBuilder();
builder.AddSheet("People", data);

// Get the Excel file as byte array
byte[] excelFile = builder.BuildFile();
```

### Advanced Example with Custom Styling

```csharp
using Gufel.ExcelBuilder;

// Create custom settings
var settings = new ExcelBuilderSettings
{
    UseDefaultStyle = true,    
    HasRowNumber = true,
    RowNumberColumnName = "Row",
    AutoFitColumns = true,
    IsRtl = true
};

// Create and configure the Excel builder
using var builder = new ExcelBuilder();
builder.SetSettings(settings);

// Create custom styles
var headerStyle = builder.CreateStyle("CustomHeader");
headerStyle.Object.Style.Fill.PatternType = ExcelFillStyle.Solid;
headerStyle.Object.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
headerStyle.Object.Style.Font.Bold = true;

var cellStyle = builder.CreateStyle("CustomCell");
cellStyle.Object.Style.Font.Name = "Arial";
cellStyle.Object.Style.Font.Size = 10;

// Add data with custom column rendering
builder.OnRenderColumn += (column, value, excelColumn, rowData) =>
{
    if (column == "Age" && value is int age)
    {
        excelColumn.Value = $"{age} years";
        return true;
    }
    return false;
};

builder.AddSheet("People", data);
byte[] excelFile = builder.BuildFile();
```

### Model with Excel Column Attributes

```csharp
public class Person
{
    [ExcelColumn(Name = "Full Name")]
    public string Name { get; set; }

    [ExcelColumn(Name = "Age", ColumnFormat = "#,##0")]
    public int Age { get; set; }

    [ExcelColumn(Name = "Birth Date", AsPersianDate = true, PersianDateFormat = "$yyyy/$MM/$dd")]
    public DateTime BirthDate { get; set; }
}
```

## Customization

### Custom Value Provider

```csharp
public class CustomValueProvider : IValueProvider
{
    public Dictionary<string, object?> GetValues(List<ExcelColumnAttribute> columns, object data)
    {
        // Implement custom value extraction logic
    }
}

// Usage
builder.SetValueProvider(new CustomValueProvider());
```

### Custom Column Provider

```csharp
public class CustomColumnProvider : IColumnProvider
{
    public List<ExcelColumnAttribute> GetColumns(Type dataType, object? data)
    {
        // Implement custom column extraction logic
    }
}

// Usage
builder.SetColumnProvider(new CustomColumnProvider());
```

## Advanced Features

### Event Handlers

The library provides several event handlers for customizing the Excel generation process:

- `OnCreateWorksheet`: Called when a new worksheet is created
- `OnCreateColumn`: Called when a new column is created
- `OnRenderColumn`: Called when a column value is being rendered

### Persian Date Support

The library includes built-in support for Persian dates with customizable formatting:

```csharp
[ExcelColumn(Name = "Birth Date", AsPersianDate = true, PersianDateFormat = "$yyyy/$MM/$dd")]
public DateTime BirthDate { get; set; }
```

### RTL Support

Enable Right-to-Left support for worksheets:

```csharp
var settings = new ExcelBuilderSettings
{
    IsRtl = true
};
```

### Memory Management

The library uses `IDisposable` pattern for proper resource management:

```csharp
using var builder = new ExcelBuilder();
// ... use the builder
// Resources will be automatically disposed
```
### Sql data reader
Add sheet with sqlreader, support custom column name and order
```csharp
using var builder = new ExcelBuilder();
var sqlConnection = new SqlConnection("sql_connection_string");
var sqlCommand = new SqlCommand("Select * from Report", sqlConnection);
sqlConnection.Open();
var reader = sqlCommand.ExecuteReader();

//this will read custom column name and order
var ctx = new List<ExcelColumnAttribute>();

var fileByte = excelBuilder.AddSheet("SqlData", reader, ctx).BuildFile();

reader.Close();
sqlCommand.Dispose();
sqlConnection.Close();
sqlConnection.Dispose();
```

## License
Please note library use epplus so please set license before using.

```csharp
 ExcelPackage.License.SetNonCommercialPersonal("Personal");
```
This project is licensed under the terms specified in the LICENSE.txt file.
