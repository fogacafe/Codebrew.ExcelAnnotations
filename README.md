# Codebrew\.ExcelAnnotations

**Codebrew\.ExcelAnnotations** is a robust and extensible C# library that streamlines the process of importing and exporting Excel files (`.xlsx`) using attribute-based configurations. Inspired by ORM frameworks like Entity Framework, it offers a fluent and customizable approach to map objects to Excel sheets and vice versa.

---

## ✨ Features

* 📅 Import Excel files directly into strongly-typed models using attribute-based mapping.
* 📄 Export models into styled Excel spreadsheets with full control via annotations and theming.
* 🎨 Custom themes and styles using abstract base classes and reflection.
* 🔄 Support for custom value converters through interfaces.
* 🧹 Extendable with your own attributes or behaviors.

---

## 📦 Installation

Install the package via NuGet:

```bash
dotnet add package Codebrew.ExcelAnnotations
```

---

## 🚀 Getting Started

### Define a model with attributes

```csharp
[WorksheetName("Product")]
public class Product
{
    [Header("Id")]
    public int Id { get; set; }

    [Header("Name", typeof(HeaderTheme))] //overwrite the default theme for header with your custom theme
    public string Name { get; set; }

    [Header("Price")]
    public decimal Price { get; set; }

    [Header("Expiration Date")]
    [DateTimeFormatter] //create custom formatters
    public DateTime? ExpirationDate { get; set; }

    [Header("Is Active")]
    [BoolConverter(["Yes"], ["No"])]
    public bool? IsActive { get; set; }

    [Header("Avalable Percentage")]
    [PercentageConverter]
    public decimal AvalablePercentage { get; set; }
}
```

### Import from Excel

```csharp
using(var importer = new Importer("excel.xlsx"))
{
    return importer.Import<T>(new WorksheetOptions()).ToList();
}
```

### Export to Excel

```csharp

using (var exporter = new Exporter())
{
    exporter.MapToWorksheet(items, new WorksheetOptions());
    exporter.MapToWorksheet(items, new WorksheetOptions() { WorksheetName = "Compare with exported" }); //can map other worksheet, forcing worksheetName on options
    exporter.ExportToPath("excel.xlsx"); // or exporter.ToStream()
}
```

---

## 🎨 Custom Themes

Create custom themes by inheriting from `ThemeBase`:

```csharp
public class MyCustomTheme : ThemeBase<MyCustomTheme>
{
    public MyCustomTheme()
    {
        FontBold = true;
        FontItalic = false;
        FontFamily = "Verdana";
        FontSize = 10;
        FontColor = ThemeColors.Text_White;
        BackgroundColor = ThemeColors.Accent1_Blue_Dark;
    }
}
```

Assign this theme in the export options to apply consistent styling.

---

## 🧐 Advanced Features

* **Dynamic Property Mapping**: Use `[Header("Column Name")]` to bind properties to spreadsheet columns.
* **Cell Value Conversion**: Implement `IConvertCellValue` to define custom parsing from cell to property.
* **Conditional Styling**: Apply styles through theming or directly via attributes.
* **Concurrent-Safe Parsing**: The engine supports safe reflection and data handling with customization flexibility.

---

## 🧪 Example Usage

Refer to the [`Example`](./Codebrew.ExcelAnnotations.Example/Program.cs) project to see the library in action with various models and customizations.

---

## 💪 Roadmap

* [ ] Localization support for headers.
* [ ] Multiple sheet exports.
* [ ] Enhanced formula and conditional formatting support.

---

## 🤝 Contributing

Contributions are welcome! Feel free to open issues or submit pull requests with improvements. Ideas for new attributes, better export settings, or utility features are always appreciated.

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 👨‍💻 Author

Developed with ☕ by [@fogacafe](https://github.com/fogacafe)
