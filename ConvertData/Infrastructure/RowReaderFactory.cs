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

            // Историческая особенность: `H2_1.xls` в проекте фактически является TSV-текстом.
            bool treatAsTsv = string.Equals(Path.GetFileName(path), "H2_1.xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".tsv", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".txt", StringComparison.OrdinalIgnoreCase);

            return treatAsTsv
                ? new TsvRowReader()
                : new EpplusRowReader();
        }
    }
}
