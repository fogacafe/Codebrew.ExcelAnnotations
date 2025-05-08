using Codebrew.ExcelAnnotations.Attributes.Base;
using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetNameAttribute : BaseAttribute
    {
        public string Name { get; }

        public SheetNameAttribute(string name, bool ignoreCase = true) : base(ignoreCase)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("Worksheet name cannot be empty", nameof(name));

            Name = name;
        }
    }
}
