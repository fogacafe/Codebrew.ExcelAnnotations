using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using Codebrew.ExcelAnnotations.Attributes.Theme;
using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class HeaderAttribute : BaseAttribute, IHeaderAttribute
    {
        public string Name { get; }
        public ITheme? Theme { get; }
        
        public HeaderAttribute(string name, Type? themeType = null) : base()
        {
            if(string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException($"{nameof(name)} shound be not empty", nameof(Name));

            Name = name;

            if (themeType is not null)
                Theme = Activator.CreateInstance(themeType) as ITheme;
            
            Theme ??= HeaderTheme.Instance;
        }

        public void ApplyStyle(IXLStyle style)
        {
            Theme?.Apply(style);
        }
    }
}
