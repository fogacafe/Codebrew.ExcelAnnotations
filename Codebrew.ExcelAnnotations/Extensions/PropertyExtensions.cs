using System;
using System.Reflection;

namespace Codebrew.ExcelAnnotations.Extensions
{
    public static class PropertyExtensions
    {
        public static bool TrySetValue(this PropertyInfo property, object obj, object value, out Exception? exception)
        {
            

            try
            {
                exception = null;
                property.SetValue(obj, value);    
                return true;
            }
            catch(Exception ex)
            {
                exception = ex;
                return false;
            }
        }
    }
}
