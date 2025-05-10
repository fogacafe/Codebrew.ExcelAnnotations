using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Engine.Interfaces;
using Codebrew.ExcelAnnotations.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            var items = new List<T>();

            foreach (var row in worksheet.RowsUsed().Skip(options.RowHeaderNumber))
            {
                var item = new T();

                foreach (var property in properties)
                {
                    if (!headers.TryGetValue(property.Name, out int columnIndex))
                        continue;

                    if (!property.TrySetValue(item, row.Cell(columnIndex).Value, out var ex) && options.ThrowWhenHasErrorToMap)
                        throw new Exception($"Error to map property '{property.Name}'. See the inner excepton for details", ex);
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
