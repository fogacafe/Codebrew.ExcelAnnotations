using System;

namespace Codebrew.ExcelAnnotations.Exceptions
{
    public class WorksheetNotFoundException : ExcelAnnotationsException
    {
        public string WorksheetName { get; }
        public WorksheetNotFoundException(string message, string worksheetName) : base(message)
        {
            WorksheetName = worksheetName;
        }
    }
}
