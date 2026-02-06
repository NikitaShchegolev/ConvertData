using System;
using System.IO;
using ConvertData.Application;

namespace ConvertData.Infrastructure
{
    internal sealed class RowReaderFactory : IRowReaderFactory
    {
        public IRowReader Create(string path)
        {
            var ext = Path.GetExtension(path);

            if (IsTextTable(ext))
                return new XlsRowReader();

            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase) && !IsRealExcelFile(path))
                return new XlsRowReader();

            return new EpplusRowReader();
        }

        private static bool IsTextTable(string? ext)
        {
            return string.Equals(ext, ".tsv", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".txt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRealExcelFile(string path)
        {
            try
            {
                Span<byte> header = stackalloc byte[8];
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                int read = fs.Read(header);
                if (read < 2)
                    return false;

                // .xlsx is ZIP
                if (header[0] == (byte)'P' && header[1] == (byte)'K')
                    return true;

                // OLE .xls
                return read >= 8
                    && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0
                    && header[4] == 0xA1 && header[5] == 0xB1 && header[6] == 0x1A && header[7] == 0xE1;
            }
            catch
            {
                return false;
            }
        }
    }
}
