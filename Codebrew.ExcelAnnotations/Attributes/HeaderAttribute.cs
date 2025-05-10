using ClosedXML.Excel;
using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class HeaderAttribute : BaseAttribute
    {
        public string Name { get; }
        public bool Bold { get; }
        public XLColor? BackColor { get; }
        public XLColor? FontColor { get; }
        public HeaderAttribute(string name, string? backColorHex = "#1d1f1e", string? fontColorHex = "#fcfcfc", bool bold = true) : base()
        {
            Name = name;
            BackColor = backColorHex != null ? XLColor.FromHtml(backColorHex) : null;
            FontColor = fontColorHex != null ? XLColor.FromHtml(fontColorHex) : null;
            Bold = bold;
        }

        public void SetStyle(IXLStyle style)
        {
            style.Font.Bold = Bold;

            if (FontColor != null)
                style.Font.FontColor = FontColor;

            if (BackColor != null)
                style.Fill.BackgroundColor = BackColor;
        }
    }
}
