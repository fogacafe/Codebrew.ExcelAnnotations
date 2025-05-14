using System;

namespace Codebrew.ExcelAnnotations.Exceptions
{
    public abstract class ExcelAnnotationsException : Exception
    {
        public ExcelAnnotationsException(string message) : base(message) { }
        public ExcelAnnotationsException(string message, Exception innerException) : base(message, innerException) { }
    }
}
