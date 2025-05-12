// See https://aka.ms/new-console-template for more information


using Codebrew.ExcelAnnotations.Engine;
using Codebrew.ExcelAnnotations.Example.Worksheets;

try
{
    var books = new List<BookWorksheet>()
    {
        new() { Title = "Apple", Date = new DateTime(2000, 01, 01), Percentage = 50m, Price = 200, IsActive = null},
        new() { Title = "DDD", Date = new DateTime(1960, 01, 01), Percentage = 100m, Price = null, IsActive = true},
        new() { Title = "Hexagonal", Date = new DateTime(2015, 01, 01), Percentage = 95m, Price = 50, IsActive = false}
    };


    var exporter = new Exporter();

    exporter.MapToWorksheet(books, new WorksheetOptions());
    exporter.MapToWorksheet(books, new WorksheetOptions() { WorksheetName = "Books25" });
    exporter.ExportToPath("excel.xlsx");
    Console.WriteLine("Finished");
}
catch(Exception ex)
{
    Console.WriteLine(ex.ToString());
}

Console.ReadLine();
