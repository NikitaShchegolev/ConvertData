using System;
using System.IO;
using ConvertData.Application;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Фабрика для создания читателей строк на основе расширения входного файла.
    /// </summary>
    internal sealed class RowReaderFactory : IRowReaderFactory
    {
        /// <summary>
        /// Создаёт соответствующий читатель строк для указанного файла.
        /// </summary>
        /// <param name="path">Путь к входному файлу.</param>
        /// <returns>Экземпляр IRowReader для чтения файла.</returns>
        /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
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
