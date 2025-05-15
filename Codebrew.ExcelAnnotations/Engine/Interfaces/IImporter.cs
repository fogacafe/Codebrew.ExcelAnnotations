using System;
using System.Collections.Generic;

namespace Codebrew.ExcelAnnotations.Engine.Interfaces
{
    public interface IImporter : IDisposable
    {
        IEnumerable<T> Import<T>(WorksheetOptions? options = null) where T : new();
    }
}
