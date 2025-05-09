using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Codebrew.ExcelAnnotations
{
    public abstract class SpreadsheetBase : IDisposable
    {
        protected readonly IXLWorkbook _workbook;

        protected SpreadsheetBase(IXLWorkbook workbook)
        {
            _workbook = workbook;
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }

        protected static string GetWorksheetName(Type type, SpreadsheetOptions options)
        {
            if (!string.IsNullOrWhiteSpace(options.SheetName))
                return options.SheetName;

            var property = type.GetCustomAttribute<SheetNameAttribute>();

            return property == null
                ? throw new ArgumentException($"Shoud use {nameof(SpreadsheetOptions)} or {nameof(SheetNameAttribute)} for set the worksheet name")
                : property.Name;
        }

        protected static List<PropertyInfo> GetProperties<T>(Type type) where T : Attribute
        {
            return type.GetProperties()
                .Where(p => p.GetCustomAttribute<T>() != null)
                .ToList();
        }
    }
}
