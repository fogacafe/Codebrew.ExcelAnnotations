using Codebrew.ExcelAnnotations.Attributes.Helpers;

namespace Codebrew.ExcelAnnotations.Attributes.Theme
{
    public class HeaderTheme : ThemeBase<HeaderTheme>
    {

        public HeaderTheme()
        {
            FontBold = true;
            FontItalic = false;
            FontFamily = "Verdana";
            FontSize = 12;
            FontColor = ThemeColors.Text_White;
            BackgroundColor = ThemeColors.Accent1_Blue_Dark;
        }
    }
}
