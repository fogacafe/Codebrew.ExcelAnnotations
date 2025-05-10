using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Engine
{
    public abstract class EngineBase
    {
        protected readonly IXLWorkbook _workbook;

        protected EngineBase(IXLWorkbook workbook)
        {
            _workbook = workbook;
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }

        protected static string GetWorksheetName(Type type, WorksheetOptions options)
        {
            if (!string.IsNullOrWhiteSpace(options.WorksheetName))
                return options.WorksheetName;

            var property = type.GetCustomAttribute<WorksheetNameAttribute>();

            return property == null
                ? throw new ArgumentException($"Shoud use {nameof(WorksheetOptions)} or {nameof(WorksheetNameAttribute)} for set the worksheet name")
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
