using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetAttribute : Attribute
    {
        public string SheetName { get; }

        public ExcelSheetAttribute(string sheetName)
        {
            SheetName = sheetName;
        }
    }
}
