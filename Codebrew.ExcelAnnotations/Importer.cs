using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations
{
    public class Importer : IDisposable
    {

        private readonly XLWorkbook _workbook;

        public Importer(XLWorkbook workbook)
        {
            _workbook = workbook;
        }

        public Importer(Stream stream)
        {
            _workbook = new XLWorkbook(stream);
        }

        public Importer(string path)
        {
            _workbook = new XLWorkbook(path);
        }

        public IEnumerable<T> Import<T>(ImporterOptions options) where T : new()
        {
            var worksheetName = GetWorksheetName(typeof(T), options);
            var worksheet = GetWorksheet(worksheetName);

            var headers = MapHeaders(worksheet, options.RowHeaderNumber);
            
            var properties = GetProperties<ExcelColumnAttribute>(typeof(T));
            var items = new List<T>();

            foreach (var row in worksheet.RowsUsed().Skip(options.RowHeaderNumber))
            {
                var item = new T();

                foreach (var property in properties)
                {
                    if (!headers.TryGetValue(property.Name, out int columnIndex))
                        continue;

                    property.TrySetValue(item, row.Cell(columnIndex).Value);
                }

                items.Add(item);
            }

            return items;
        }

        private Dictionary<string, int> MapHeaders(IXLWorksheet sheet, int rowHeaderNumber = 1)
        {
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            var rowHeader = sheet
                .Rows()
                .Where(x => x.RowNumber() == rowHeaderNumber)
                .FirstOrDefault();

            foreach (var cell in rowHeader.CellsUsed())
                headers.Add(cell.GetString(), cell.Address.ColumnNumber);

            return headers;
        }

        private static List<PropertyInfo> GetProperties<T>(Type type) where T : Attribute
        {
            return type.GetProperties()
                .Where(p => p.GetCustomAttribute<T>() != null)
                .ToList();
        }

        private static string GetWorksheetName(Type type, ImporterOptions options)
        {
            if (!string.IsNullOrWhiteSpace(options.SheetName))
                return options.SheetName;

            var property = type.GetCustomAttribute<SheetNameAttribute>();

            return property == null
                ? throw new ArgumentException($"Shoud use {nameof(ImporterOptions)} or {nameof(SheetNameAttribute)} for set the worksheet name")
                : property.Name;
        }

        private IXLWorksheet GetWorksheet(string worksheetName)
        {
            var worksheet = _workbook.Worksheets
                .Where(x => x.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (worksheet == null)
                throw new NullReferenceException($"Sheet with name {worksheetName} was not found");

            return worksheet;
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }
    }
}
