using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using Codebrew.ExcelAnnotations.Engine.Interfaces;
using Codebrew.ExcelAnnotations.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Engine
{
    public class Importer : EngineBase, IImporter
    {

        public Importer(IXLWorkbook workbook) : base(workbook) { }

        public Importer(Stream stream) : base(new XLWorkbook(stream)) { }

        public Importer(string path) : base(new XLWorkbook(path)) { }

        public IEnumerable<T> Import<T>(WorksheetOptions options) where T : new()
        {
            var worksheetName = GetWorksheetName(typeof(T), options);
            var worksheet = GetWorksheet(worksheetName);

            var headers = MapHeaders(worksheet, options.RowHeaderNumber);

            var properties = GetProperties<HeaderAttribute>(typeof(T));

            var propConfigs = properties.Select((property, index) =>
            {
                var converter = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IConvertCellValue>()
                    .FirstOrDefault();

                var column = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IHeaderAttribute>()
                    .FirstOrDefault();

                return new
                {
                    Index = index,
                    Converter = converter,
                    Column = column,
                    Property = property
                };

            }).ToList();


            var items = new List<T>();

            foreach (var row in worksheet.RowsUsed().Skip(options.RowHeaderNumber))
            {
                var item = new T();

                foreach(var config in propConfigs)
                {
                    if (!headers.TryGetValue(config.Column.Name, out var columnIndex))
                        continue;

                    Exception? ex = null;

                    if (config.Converter != null)
                        config.Property.TrySetValue(item, config.Converter.FromCellValue(row.Cell(columnIndex)), out ex);
                    else
                        config.Property.TrySetValue(item, row.Cell(columnIndex).Value, out ex);

                    if (ex != null && options.ThrowWhenHasErrorToMap)
                        throw new Exception($"Error to map property '{config.Property.Name}'. See the inner excepton for details", ex);

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

        private IXLWorksheet GetWorksheet(string worksheetName)
        {
            var worksheet = _workbook.Worksheets
                .Where(x => x.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (worksheet == null)
                throw new NullReferenceException($"Sheet with name {worksheetName} was not found");

            return worksheet;
        }
    }
}
