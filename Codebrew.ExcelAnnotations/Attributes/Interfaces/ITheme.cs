using ClosedXML.Excel;

namespace Codebrew.ExcelAnnotations.Attributes.Interfaces
{
    public interface ITheme
    {
        void Apply(IXLStyle style);
    }
}
