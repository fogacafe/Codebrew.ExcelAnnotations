using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Base;
using Codebrew.ExcelAnnotations.Attributes.Converters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations
{
    public class Exporter : SpreadsheetBase
    {
        public Exporter() : base(new XLWorkbook()) { }

        public Exporter(IXLWorkbook workbook) : base(workbook) { }


        public void Export<T>(IEnumerable<T> items, SpreadsheetOptions options)
        {
            var worksheetName = GetWorksheetName(typeof(T), options);
            var worksheet = _workbook.AddWorksheet(worksheetName);

            var properties = GetProperties<ColumnNameAttribute>(typeof(T));

            var propConfigs = properties.Select((property, index) =>
            {
                var converter = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IConvertToCell>()
                    .FirstOrDefault();

                var columnName = property.GetCustomAttribute<ColumnNameAttribute>().Name;

                return new
                {
                    Index = index,
                    Converter = converter,
                    ColumnName = columnName,
                    Property = property
                };
            }).ToList();

            var rowNumber = options.RowHeaderNumber;
            foreach (var item in items)
            {
                var row = worksheet.Row(rowNumber++);

                foreach (var config in propConfigs)
                {
                    var value = config.Property.GetValue(item, null);
                    var cell = row.Cell(config.Index + 1);

                    cell.Value = config.Converter != null
                        ? config.Converter.ToXLCellValue(value)
                        : cell.Value = ConvertToXLCellValue(value);
                }
            }
        }

        protected XLCellValue ConvertToXLCellValue(object? value)
        {
            return value switch
            {
                null => Blank.Value,
                string s => s,
                double d => d,
                float f => f,
                int i => i,
                long l => l,
                decimal dec => dec,
                bool b => b ? "Yes" : "No",
                DateTime dt => dt,
                TimeSpan ts => ts,
                Enum e => e.ToString(),
                Guid g => g.ToString(),
                _ => value?.ToString() ?? string.Empty
            };
        }

        public Stream ExportToStream()
        {
            var memoryStream = new MemoryStream();
            _workbook.SaveAs(memoryStream);
            return memoryStream;
        }

        public void ExportToPath<T>(string path)
        {
            _workbook.SaveAs(path);
        }
    }
}
