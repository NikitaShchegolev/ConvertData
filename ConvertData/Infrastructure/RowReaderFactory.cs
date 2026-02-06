using System;
using System.IO;
using ConvertData.Application;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Фабрика reader-ов.
    /// Выбирает, как читать конкретный файл: как TSV (для текстовых табличных данных) или как Excel через EPPlus.
    /// </summary>
    internal sealed class RowReaderFactory : IRowReaderFactory
    {
        /// <summary>
        /// Возвращает реализацию `IRowReader`, подходящую для переданного файла.
        /// </summary>
        /// <param name="path">Путь к входному файлу.</param>
        /// <returns>Экземпляр reader-а.</returns>
        public IRowReader Create(string path)
        {
            var ext = Path.GetExtension(path);

            // Текстовые форматы по расширению.
            if (string.Equals(ext, ".tsv", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".txt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase))
            {
                return new XlsRowReader();
            }

            // Некоторые входные файлы имеют расширение .xls, но фактически являются TSV.
            // Пытаемся определить это по содержимому (сигнатуры Excel + наличие табов в первых байтах).
            if (LooksLikeTsvWithXlsExtension(path))
                return new XlsRowReader();

            return new EpplusRowReader();
        }

        private static bool LooksLikeTsvWithXlsExtension(string path)
        {
            var ext = Path.GetExtension(path);
            if (!string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase))
                return false;

            try
            {
                Span<byte> header = stackalloc byte[512];
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                int read = fs.Read(header);
                if (read <= 0)
                    return false;

                // Настоящий .xlsx начинается с PK
                if (read >= 2 && header[0] == (byte)'P' && header[1] == (byte)'K')
                    return false;

                // Настоящий OLE .xls начинается с D0 CF 11 E0 A1 B1 1A E1
                if (read >= 8
                    && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0
                    && header[4] == 0xA1 && header[5] == 0xB1 && header[6] == 0x1A && header[7] == 0xE1)
                    return false;

                // Если в первых байтах много табов/переводов строк — это почти наверняка TSV.
                int tabs = 0;
                int newlines = 0;
                int binary = 0;

                for (int i = 0; i < read; i++)
                {
                    byte b = header[i];
                    if (b == (byte)'\t') tabs++;
                    else if (b == (byte)'\n' || b == (byte)'\r') newlines++;
                    else if (b == 0) binary++;
                }

                if (binary > 0)
                    return false;

                return tabs >= 2 && newlines >= 1;
            }
            catch
            {
                return false;
            }
        }
    }
}
