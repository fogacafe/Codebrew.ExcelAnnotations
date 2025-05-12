using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using System;
using System.Linq;

namespace Codebrew.ExcelAnnotations.Attributes.Converters
{
    [AttributeUsage(AttributeTargets.Property)]
    public class BoolConverterAttribute : BaseAttribute, IConvertCellValue
    {
        public string[] TrueValues { get; }
        public string[] FalseValues { get; }
        
        public BoolConverterAttribute(string[] trueValues,
                                      string[] falseValues) : base()
        {
            ValidateArray(trueValues, nameof(trueValues));
            ValidateArray(falseValues, nameof(falseValues));

            TrueValues = trueValues;
            FalseValues = falseValues;
        }

        public string TrueValue => TrueValues[0];
        public string FalseValue => FalseValues[0];

        private void ValidateArray(string[] arr, string propertyName)
        {
            if (arr is null || arr.Length == 0 || arr.All(x => string.IsNullOrWhiteSpace(x)))
                throw new ArgumentException($"{propertyName} must have a value.", propertyName);
        }

        public XLCellValue ToCellValue(object? value)
        {
            if (value is null)
                return Blank.Value;

            return (bool)value ? TrueValue :FalseValue;
        }

        public object? FromCellValue(IXLCell cell)
        {
            if (cell.TryGetValue<string>(out var text) && !string.IsNullOrWhiteSpace(text))
                return TrueValues.Contains(text, StringComparer);
            
            return null;
        }
    }
}

