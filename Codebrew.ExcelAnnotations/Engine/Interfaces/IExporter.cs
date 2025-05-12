using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Codebrew.ExcelAnnotations.Engine.Interfaces
{
    public interface IExporter : IDisposable
    {
        void MapToWorksheet<T>(IEnumerable<T> items, WorksheetOptions options);
        public Stream ExportToStream();

        public void ExportToPath(string path);
    }
}
