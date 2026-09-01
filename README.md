# Codebrew.ExcelAnnotations

> **Deprecated / superseded.** This project has been replaced by **[Indtec.ExcelMapper](https://github.com/fogacafe/indtec-labz-excel)**, a redesigned and actively maintained Excel mapper for .NET.

`Codebrew.ExcelAnnotations` was the original experiment behind the idea: map strongly typed C# models to Excel using annotations, converters and reusable styling. It proved useful in practice and eventually evolved into a new implementation focused on stronger typing, cleaner abstractions and modern .NET tooling.

## Use Indtec.ExcelMapper instead

```bash
dotnet add package Indtec.ExcelMapper
```

The successor keeps the original goal, but improves the architecture substantially:

- Roslyn incremental source generator instead of runtime property discovery.
- Strongly typed generated getters and setters.
- `netstandard2.0` and `net8.0` support.
- Required-column validation.
- Custom value converters without leaking ClosedXML types.
- Reusable typed themes.
- Row-aware conditional styling with cross-column rules.
- CI, tests and automated NuGet publishing through GitHub OIDC.

Example:

```csharp
[ExcelSheet("Products")]
public partial class Product
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Cost", Order = 3)]
    public decimal Cost { get; set; }

    [ExcelColumn("Price", Order = 4)]
    public decimal Price { get; set; }
}
```

```csharp
var mapper = new ExcelMapper();

mapper.Export(products, "products.xlsx", options =>
{
    options.Column(x => x.Price)
        .When(row => row.Price < row.Cost)
        .Background("#FFCCCC")
        .Bold();
});
```

## Historical package

The old NuGet package remains available for existing consumers, but new projects should use `Indtec.ExcelMapper`.

This repository is kept as the historical predecessor of the current library rather than being deleted.

## License

MIT
