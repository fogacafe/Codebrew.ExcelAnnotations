using System.Reflection;

namespace Codebrew.ExcelAnnotations.Extensions
{
    public static class PropertyExtensions
    {
        public static bool TrySetValue(this PropertyInfo property, object obj, object value)
        {
            try
            {
                property.SetValue(obj, value);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
