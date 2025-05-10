using ClosedXML.Excel;

namespace Codebrew.ExcelAnnotations.Attributes.Converters
{
    public interface IConvertToCell
    {
        XLCellValue ToXLCellValue(object value);
    }
}
