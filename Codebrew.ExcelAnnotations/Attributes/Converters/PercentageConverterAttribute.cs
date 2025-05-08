using Codebrew.ExcelAnnotations.Attributes.Base;
using System;
using System.Globalization;

namespace Codebrew.ExcelAnnotations.Attributes.Converters
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PercentageConverterAttribute : BaseAttribute
    {
        private readonly CultureInfo _culture;
        public PercentageConverterAttribute(string culture) : base()
        {
            _culture = new CultureInfo(culture);
        }

        public PercentageConverterAttribute() : base()
        {
            _culture = new CultureInfo("en-US");
        }

        public decimal? Convert(string value)
        {
            if(string.IsNullOrWhiteSpace(value)) return null;

            bool isPercentageformat = false;
            if (value.Contains("%"))
            {
                value = value.Replace("%", "");
                isPercentageformat = true;
            }

            value = value.Trim();

            if(decimal.TryParse(value, NumberStyles.Any, _culture, out decimal percentage))
                return isPercentageformat ? percentage : percentage * 100;

            return 0;
        }
    }
}
