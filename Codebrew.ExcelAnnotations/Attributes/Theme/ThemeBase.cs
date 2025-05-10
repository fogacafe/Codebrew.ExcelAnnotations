using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes.Helpers;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;

namespace Codebrew.ExcelAnnotations.Attributes.Theme
{
    public abstract class ThemeBase<T> : ITheme
        where T : ITheme, new()
    {
        public bool FontBold { get; protected set; }
        public bool FontItalic { get; protected set; }
        public string FontFamily { get; protected set; }
        public int FontSize { get; protected set; }
        public string FontColor { get; protected set; }
        public string BackgroundColor { get; protected set; }

        private static ITheme? _instance;
        public static ITheme Instance
        {
            get
            {
                _instance ??= new T();
                return _instance;
            }
        }

        protected ThemeBase()
        {
            FontBold = false;
            FontItalic = false;
            FontFamily = "Verdana";
            FontSize = 10;
            FontColor = ThemeColors.Text_Black;
            BackgroundColor = string.Empty;
        }

        public void Apply(IXLStyle style)
        {
            style.Font.FontName = FontFamily;
            style.Font.FontSize = FontSize;
            style.Font.Bold = FontBold;
            style.Font.Italic = FontItalic;
            style.Fill.BackgroundColor = XLColor.FromHtml(BackgroundColor);
            style.Font.FontColor = XLColor.FromHtml(FontColor);
        }
    }
}
