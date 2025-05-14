using Codebrew.ExcelAnnotations.Attributes.Interfaces;
using System;
using System.Collections.Concurrent;

namespace Codebrew.ExcelAnnotations.Attributes.Helpers
{
    public static class ThemeHelper
    {
        private static ConcurrentDictionary<Type, ITheme> _themes = new();

        public static ITheme GetTheme<T>() where T : ITheme, new()
        {
            return GetTheme(typeof(T));
        }

        public static ITheme GetTheme(Type themeType)
        {
            if (!typeof(ITheme).IsAssignableFrom(themeType))
                throw new ArgumentException($"{nameof(themeType)} should be instance of type {nameof(ITheme)}", nameof(themeType));

            return _themes.GetOrAdd(themeType, type =>
            {
                var theme = Activator.CreateInstance(type) as ITheme;
                return theme!;
            });
        }
    }
}
