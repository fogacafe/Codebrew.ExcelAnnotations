using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using Codebrew.ExcelAnnotations.Engine.Interfaces;
using Codebrew.ExcelAnnotations.Exceptions;
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
            var propConfigs = MapPropertiesConfig<T>();

            var items = new List<T>();
            var rows = worksheet.RowsUsed().Skip(options.RowHeaderNumber);

            foreach (var row in rows)
            {
                var item = new T();

                foreach(var config in propConfigs)
                {
                    if (!headers.TryGetValue(config.Header.Name, out var columnIndex))
                        continue;

                    Exception? ex = null;

                    var cell = row.Cell(columnIndex);

                    if (config.Converter != null)
                        config.Property.TrySetValue(item, config.Converter.FromCellValue(cell), out ex);
                    else
                        config.Property.TrySetValue(item, cell.Value, out ex);

                    if (ex != null && options.ThrowWhenHasErrorToMap)
                    {
                        throw new SetPropertyValueException($"Error to map property '{config.Property.Name} at column '{config.Header.Name}'. See the inner excepton for details", 
                            config.Property.Name, 
                            config.Header.Name, 
                            ex);
                    }
                        

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

            var cellsUsed = rowHeader.CellsUsed();
            foreach (var cell in cellsUsed)
                headers.Add(cell.GetString(), cell.Address.ColumnNumber);

            return headers;
        }

        private IXLWorksheet GetWorksheet(string worksheetName)
        {
            var worksheet = _workbook.Worksheets
                .Where(x => x.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (worksheet == null)
                throw new WorksheetNotFoundException($"Worksheet with name {worksheetName} was not found", worksheetName);

            return worksheet;
        }

        private List<PropertyMapConfig> MapPropertiesConfig<T>()
        {
            List<PropertyMapConfig> propConfigs = typeof(T).GetProperties()
                .Select((property, index) =>
                {
                    var header = property.GetCustomAttributes<BaseAttribute>(true)
                        .OfType<IHeaderAttribute>()
                        .FirstOrDefault();

                    if (header is null)
                        return null;

                    var converter = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IConvertCellValue>()
                    .FirstOrDefault();

                    return new PropertyMapConfig(
                        index,
                        property,
                        converter,
                        header
                    );
                }).Where(x => x != null)
                .ToList()!;

            return propConfigs;
        }
    }
}
