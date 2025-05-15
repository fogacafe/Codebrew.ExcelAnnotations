using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using Codebrew.ExcelAnnotations.Exceptions;
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
            if (workbook is null)
                throw new NullReferenceException($"{nameof(workbook)} is null");

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
                ? throw new WorksheetNameNotInformedException($"Shoud use '{nameof(WorksheetOptions)}' or '{nameof(WorksheetNameAttribute)}' for set the worksheet name")
                : property.Name;
        }

        protected static List<PropertyMapConfig> MapPropertiesConfig<T>()
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

                    var styler = property.GetCustomAttributes<BaseAttribute>(true)
                   .OfType<IStylerCell>()
                   .FirstOrDefault();

                    return new PropertyMapConfig(
                        index,
                        property,
                        converter,
                        header,
                        styler
                    );
                }).Where(x => x != null)
                .ToList()!;

            return propConfigs;
        }
    }
}
