using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using Codebrew.ExcelAnnotations.Engine.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Engine
{
    public class Exporter : EngineBase, IExporter
    {
        public Exporter() : base(new XLWorkbook()) { }

        public Exporter(IXLWorkbook workbook) : base(workbook) { }


        public void MapToWorksheet<T>(IEnumerable<T> items, WorksheetOptions? options)
        {
            options ??= new WorksheetOptions();

            var worksheetName = GetWorksheetName(typeof(T), options);
            var worksheet = _workbook.AddWorksheet(worksheetName);

            var properties = GetProperties<HeaderAttribute>(typeof(T));

            var propConfigs = properties.Select((property, index) =>
            {
                var converter = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IConvertCellValue>()
                    .FirstOrDefault();

                var styler = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IStylerCell>()
                    .FirstOrDefault();

                var column = property.GetCustomAttributes<BaseAttribute>(true)
                    .OfType<IHeaderAttribute>()
                    .FirstOrDefault();

                return new
                {
                    Index = index,
                    Converter = converter,
                    Column = column,
                    Property = property,
                    Styler = styler
                };
            }).ToList();

            var rowNumber = options.RowHeaderNumber + 1;
            foreach (var item in items)
            {
                var row = worksheet.Row(rowNumber++);

                foreach (var config in propConfigs)
                {
                    var value = config.Property.GetValue(item, null);
                    var cell = row.Cell(config.Index + 1);

                    cell.Value = config.Converter != null
                        ? config.Converter.ToCellValue(value)
                        : cell.Value = ConvertToXLCellValue(value);

                    config.Styler?.Apply(cell.Style);
                }
            }

            var headerRow = worksheet.Row(options.RowHeaderNumber);
            foreach (var propConfig in propConfigs)
            {
                var cell = headerRow.Cell(propConfig.Index + 1);
                cell.Value = propConfig.Column.Name;
                propConfig.Column.ApplyStyle(cell.Style);

                worksheet.Column(propConfig.Index + 1).AdjustToContents();
            }
        }

        protected virtual XLCellValue ConvertToXLCellValue(object? value)
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

        public void ExportToPath(string path)
        {
            _workbook.SaveAs(path);
        }
    }
}
