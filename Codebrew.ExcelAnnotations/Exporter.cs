using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations
{
    public static class Exporter
    {
        public static Stream Export<T>(IEnumerable<T> data)
        {
            var workbook = new XLWorkbook();
            var type = typeof(T);

            var worksheet = workbook.Worksheets.Add(GetSheetName(type));

            var props = GetProperties<ExcelColumnAttribute>(type);

            MapHeader(worksheet, data, props);
            MapRows(worksheet, data, props);
            
            var memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);

            return memoryStream;
        }

        private static void MapHeader<T>(IXLWorksheet worksheet, IEnumerable<T> data, List<PropertyInfo> props)
        {
            for (var index = 0; index < props.Count; index++)
            {
                var columnAttibute = props[index].GetCustomAttribute<ExcelColumnAttribute>();
                worksheet.Cell(1, index + 1).Value = columnAttibute.ColumnName;
            }
        }

        private static void MapRows<T>(IXLWorksheet worksheet, IEnumerable<T> data, List<PropertyInfo> props)
        {
            int row = 2;
            foreach (var item in data)
            {
                for (int index = 0; index < props.Count; index++)
                {
                    var value = props[index].GetValue(item);
                    worksheet.Cell(row, index + 1).Value = value switch
                    {
                        null => Blank.Value,
                        string s => s,
                        int i => i,
                        long l => l,
                        short s => s,
                        float f => f,
                        double d => d,
                        decimal m => (double)m,
                        DateTime dt => dt,
                        bool b => b,
                        TimeSpan ts => ts,
                        _ => value.ToString()
                    };
                }
                row++;
            }
        }

        private static List<PropertyInfo> GetProperties<T>(Type type) where T : Attribute
        {
            return type.GetProperties()
                .Where(p => p.GetCustomAttribute<T>() != null)
                .ToList();
        }

        private static string GetSheetName(Type type)
        {
            var sheetAttribute = type.GetCustomAttribute<ExcelSheetAttribute>();
            return sheetAttribute?.SheetName ?? "Sheet1";
        }
    }
}
