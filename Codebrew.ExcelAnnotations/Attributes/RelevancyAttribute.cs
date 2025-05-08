using Codebrew.ExcelAnnotations.Attributes.Base;
using Codebrew.ExcelAnnotations.Attributes.Enums;
using System;

namespace Codebrew.ExcelAnnotations.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class RelevancyAttribute : BaseAttribute
    {
        public Relevancy Relevancy { get; }

        public RelevancyAttribute(Relevancy relevancy)
        {
            Relevancy = relevancy;
        }
    }
}
