using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using System;

namespace Codebrew.ExcelAnnotations.Attributes.Formatters
{
    public class DateTimeFormatterAttribute : Attribute, IStylerCell
    {
        private string _format;

        public DateTimeFormatterAttribute(string format = "dd/mm/yyyy")
        {
            if(string.IsNullOrWhiteSpace(format))
                throw new ArgumentNullException($"{nameof(format)} should not be empty", nameof(format));

            _format = format;
        }

        public void Apply(IXLStyle style)
        {
            style.DateFormat.SetFormat(_format);
        }
    }
}
