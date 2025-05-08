using Codebrew.ExcelAnnotations.Attributes.Base;
using System;
using System.Linq;

namespace Codebrew.ExcelAnnotations.Attributes.Converters
{
    [AttributeUsage(AttributeTargets.Property)]
    public class BoolConverterAttribute : BaseAttribute
    {
        public string[] TrueValues { get; }
        public string[] FalseValues { get; }
        
        public BoolConverterAttribute(string[] trueValues,
                                      string[] falseValues,
                                      bool ignoreCase = true) : base(ignoreCase)
        {
            ValidateArray(trueValues, nameof(trueValues));
            ValidateArray(falseValues, nameof(falseValues));

            TrueValues = trueValues;
            FalseValues = falseValues;
        }

        public string TrueValue => TrueValues[0];
        public string FalseValue => FalseValues[0];

        public bool Convert(string value)
        {
            return !string.IsNullOrWhiteSpace(value) 
                && TrueValues.Contains(value, StringComparer);
        }

        public string Convert(bool value)
        {
            return value ? TrueValue : FalseValue;
        }

        private void ValidateArray(string[] arr, string propertyName)
        {
            if (arr is null || arr.Length == 0 || arr.All(x => string.IsNullOrWhiteSpace(x)))
                throw new ArgumentException($"{propertyName} must have a value.", propertyName);
        }
    }
}

