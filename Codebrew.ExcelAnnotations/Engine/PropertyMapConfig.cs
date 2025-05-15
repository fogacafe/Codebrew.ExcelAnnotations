using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Engine
{
    public record PropertyMapConfig
    {
        public PropertyMapConfig(int index,
                                 PropertyInfo property,
                                 IConvertCellValue? converter,
                                 IHeaderAttribute header,
                                 IStylerCell styler)
        {
            Index = index;
            Property = property;
            Converter = converter;
            Header = header;
            Styler = styler;
        }

        public int Index { get; private set; }
        public PropertyInfo Property { get; private set; } = null!;
        public IConvertCellValue? Converter { get; private set; }
        public IHeaderAttribute Header { get; private set; } = null!;
        public IStylerCell Styler { get; private set; }
    }
    
}
