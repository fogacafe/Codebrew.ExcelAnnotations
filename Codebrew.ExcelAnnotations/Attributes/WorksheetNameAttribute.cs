using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class WorksheetNameAttribute : BaseAttribute
    {
        public string Name { get; }

        public WorksheetNameAttribute(string name, bool ignoreCase = true) : base(ignoreCase)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException("Worksheet name cannot be empty", nameof(name));

            Name = name;
        }
    }
}
