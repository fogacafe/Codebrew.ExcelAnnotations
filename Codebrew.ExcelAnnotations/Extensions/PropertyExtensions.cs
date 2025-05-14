using ClosedXML.Excel;
using System;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Extensions
{
    public static class PropertyExtensions
    {
        public static bool TrySetValue(this PropertyInfo property, object obj, object? value, out Exception? exception)
        {
            try
            {
                exception = null;

                var propertyType = property.PropertyType;
                var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

                var convertedValue = ConvertValue(value, targetType);

                if (convertedValue is null)
                {
                    var finalValue = AcceptsNull(propertyType)
                        ? null
                        : GetDefaultValue(targetType);

                    property.SetValue(obj, finalValue);
                }
                else
                {
                    property.SetValue(obj, convertedValue);
                }

                return true;
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
        }

        private static object? ConvertValue(object? value, Type targetType)
        {
            if (value is XLCellValue xlVal)
            {
                object? raw = xlVal.IsBlank ? null
                              : xlVal.IsDateTime ? xlVal.GetDateTime()
                              : xlVal.IsNumber ? (decimal)xlVal.GetNumber()
                              : xlVal.IsText ? xlVal.GetText()
                              : xlVal.IsBoolean ? xlVal.GetBoolean()
                              : (object?)xlVal;

                if (raw == null)
                    return null;

                if (targetType.IsEnum && raw is string str)
                    return Enum.Parse(targetType, str, ignoreCase: true);

                return Convert.ChangeType(raw, targetType);
            }

            if (value == null)
                return null;

            return Convert.ChangeType(value, targetType);
        }

        private static bool AcceptsNull(Type type)
        {
            return !type.IsValueType || Nullable.GetUnderlyingType(type) != null;
        }

        private static object? GetDefaultValue(Type type)
        {
            return !type.IsValueType || Nullable.GetUnderlyingType(type) != null
                ? null
                : Activator.CreateInstance(type);
        }
    }
}
