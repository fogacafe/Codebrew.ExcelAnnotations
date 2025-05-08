using System;

namespace Codebrew.ExcelAnnotations.Attributes.Base
{
    public abstract class BaseAttribute : Attribute
    {
        protected bool IgnoreCase { get; }
        protected BaseAttribute(bool ignoreCase = true)
        {
            IgnoreCase = ignoreCase;
        }

        protected StringComparer StringComparer 
            => IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;

        protected StringComparison StringComparison
            => IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    }
}
