// See https://aka.ms/new-console-template for more information


using Codebrew.ExcelAnnotations.Engine;
using Codebrew.ExcelAnnotations.Example.Worksheets;

try
{
    Console.WriteLine("Creating");
    Create();
    Console.WriteLine("Created");

    Console.WriteLine("Importing");
    var items = Import<BookWorksheet>();
    Console.WriteLine("Imported");

    Console.WriteLine("Exporting");
    Export(items);
    Console.WriteLine("Exported");


    Console.WriteLine("Finished");
    Console.ReadLine();
}
catch(Exception ex)
{
    Console.WriteLine(ex.ToString());
}


static List<T> Import<T>() where T : new()
{
    using(var importer = new Importer("excel.xlsx"))
    {
        return importer.Import<T>(new WorksheetOptions()).ToList();
    }
}

static void Create()
{
    var books = new List<BookWorksheet>()
    {
        new() { Title = "Apple", Date = new DateTime(2000, 01, 01), Percentage = 50m, Price = 200, IsActive = null},
        new() { Title = "DDD", Date = new DateTime(1960, 01, 01), Percentage = 100m, Price = null, IsActive = true},
        new() { Title = "Hexagonal", Date = new DateTime(2015, 01, 01), Percentage = 95m, Price = 50, IsActive = false}
    };

    Export(books);
}

static void Export<T>(IEnumerable<T> items)
{
    using (var exporter = new Exporter())
    {
        exporter.MapToWorksheet(items, new WorksheetOptions());
        exporter.MapToWorksheet(items, new WorksheetOptions() { WorksheetName = "Books25" });
        exporter.ExportToPath("excel.xlsx");
    }
}