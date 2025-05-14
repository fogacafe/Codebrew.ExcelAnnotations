using System;

namespace Codebrew.ExcelAnnotations.Exceptions
{
    public class SetPropertyValueException : ExcelAnnotationsException
    {
        public string PropertyName { get; }
        public string HeaderName { get; }
        public SetPropertyValueException(string message, string propertyName, string headerName) : base(message)
        {
            PropertyName = propertyName;
            HeaderName = headerName;
        }

        public SetPropertyValueException(string message, string propertyName, string headerName, Exception innerException) : base(message, innerException)
        {
            PropertyName = propertyName;
            HeaderName = headerName;
        }
    }
}
