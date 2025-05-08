using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Codebrew.ExcelAnnotations
{
    public class Any
    {
        private readonly object? _value;

        public Any(object? value)
        {
            _value = value;
        }

        public static implicit operator Any(DateTime value) => new Any(value);
        public static implicit operator Any(string value) => new Any(value);
        public static implicit operator Any(double value) => new Any(value);
        public static implicit operator Any(int value) => new Any(value);
        public static implicit operator Any(bool value) => new Any(value);
        public static implicit operator Any(decimal value) => new Any(value);
        public static implicit operator Any(long value) => new Any(value);
        public static implicit operator Any(float value) => new Any(value);
        public static implicit operator Any(DateTime? value) => new Any(value);
        public static implicit operator Any(double? value) => new Any(value);
        public static implicit operator Any(int? value) => new Any(value);
        public static implicit operator Any(bool? value) => new Any(value);
        public static implicit operator Any(decimal? value) => new Any(value);
        public static implicit operator Any(long? value) => new Any(value);
        public static implicit operator Any(float? value) => new Any(value);


        public static implicit operator string(Any val) => ConvertTo<string>(val._value);
        public static implicit operator DateTime(Any val) => ConvertTo<DateTime>(val._value);
        public static implicit operator int(Any val) => ConvertTo<int>(val._value);
        public static implicit operator double(Any val) => ConvertTo<double>(val._value);
        public static implicit operator bool(Any val) => ConvertTo<bool>(val._value);
        public static implicit operator decimal(Any val) => ConvertTo<decimal>(val._value);
        public static implicit operator long(Any val) => ConvertTo<long>(val._value);
        public static implicit operator float(Any val) => ConvertTo<float>(val._value);
        public static implicit operator DateTime?(Any val) => ConvertTo<DateTime>(val._value);
        public static implicit operator int?(Any val) => ConvertTo<int>(val._value);
        public static implicit operator double?(Any val) => ConvertTo<double>(val._value);
        public static implicit operator bool?(Any val) => ConvertTo<bool>(val._value);
        public static implicit operator decimal?(Any val) => ConvertTo<decimal>(val._value);
        public static implicit operator long?(Any val) => ConvertTo<long>(val._value);
        public static implicit operator float?(Any val) => ConvertTo<float>(val._value);

        public static implicit operator XLCellValue(Any v)
        {
            return v;
        }

        private static T ConvertTo<T>(object? value)
        {
            if (value == null || value is DBNull)
                return default!;

            if (value is T t)
                return t;

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch
            {
                throw new InvalidCastException($"Cannot convert value '{value}' to type {typeof(T).Name}");
            }
        }
    }
}
