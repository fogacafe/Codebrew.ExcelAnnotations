using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using System.Collections.Generic;
using System.IO;
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



            var rowNumber = options.RowHeaderNumber + 1;
            foreach (var item in items)
            {
                var row = worksheet.Row(rowNumber);

                for (int index = 1; index <= properties.Count; index++)
                {
                    var columnAtt = properties[index].GetCustomAttribute<ColumnNameAttribute>();
                    row.Cell(index + 1).Value = properties[index].GetValue(item, null)?.ToString();
                }
            }
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
