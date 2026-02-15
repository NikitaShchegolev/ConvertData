using System;
using System.IO;
using ConvertData.Application;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure
{
    internal sealed class RowReaderFactory : IRowReaderFactory
    {
        public IRowReader Create(string path)
        {
            var ext = Path.GetExtension(path);

            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                _ = ExcelFileSignature.Detect(path);
                return new EpplusRowReader();
            }

            throw new NotSupportedException("Only .xls/.xlsx inputs are supported.");
        }
    }
}
