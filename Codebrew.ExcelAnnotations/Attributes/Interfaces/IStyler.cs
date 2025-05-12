using ClosedXML.Excel;

namespace Codebrew.ExcelAnnotations.Attributes.Interfaces
{
    public interface IStyler
    {
        void ApplyStyle(IXLStyle style);
    }
}
