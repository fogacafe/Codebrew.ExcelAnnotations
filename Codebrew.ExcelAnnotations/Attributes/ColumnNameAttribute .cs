using System;
using Codebrew.ExcelAnnotations.Attributes.Base;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : BaseAttribute
    {
        public string Name { get; }
        private bool IsRequired { get; set; }
        public ExcelColumnAttribute(string name, bool isRequired = false, bool ignoreCase = true) : base(ignoreCase)
        {
            Name = name;
            IsRequired = isRequired;
        }
    }
}
