namespace Codebrew.ExcelAnnotations
{
    public class ImporterOptions
    {
        /// <summary>
        /// SheetName
        /// if empty, use the SheetNameAttribute
        /// </summary>
        public string SheetName { get; set; } = string.Empty;

        /// <summary>
        /// RowHeaderNumber
        /// Default is 1
        /// </summary>
        public int RowHeaderNumber { get; set; } = 1;

        /// <summary>
        /// Throw exception when has error to map a property
        /// Defaul is false
        /// </summary>
        public bool ThrowWhenHasErrorToMap { get; set; } = false;
    }
}
