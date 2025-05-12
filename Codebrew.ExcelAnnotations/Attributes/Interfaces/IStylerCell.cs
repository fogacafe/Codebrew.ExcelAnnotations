using ClosedXML.Excel;

namespace Codebrew.ExcelAnnotations.Attributes.Interfaces
{
    public interface IStylerCell
    {
        void Apply(IXLStyle style);
    }
}
