using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using System;
using System.Globalization;

namespace Codebrew.ExcelAnnotations.Attributes.Converters
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PercentageConverterAttribute : BaseAttribute, IConvertCellValue, IStylerCell
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

        public void Apply(IXLStyle style)
        {
            style.NumberFormat.Format = "0.00%";
        }

        public XLCellValue ToCellValue(object? value)
        {
            decimal? percentage = (decimal?)value;
            return percentage is null ? null : percentage / 100;
        }

        public object? FromCellValue(IXLCell cell)
        {
            if (cell.TryGetValue<decimal?>(out var value) && value != null)
                return value * 100;

            return null;
        }
    }
}
