using ClosedXML.Excel;

namespace Codebrew.ExcelAnnotations.Attributes.Interfaces
{
    public interface IConvertCellValue
    {
        XLCellValue ToCellValue(object? value);
        object? FromCellValue(IXLCell cell);
    }
}
