using System.Collections.Generic;
using ConvertData.Domain;

namespace ConvertData.Application
{
    /// <summary>
    /// Контракт сохранения набора строк `Row` в целевой формат (например JSON).
    /// </summary>
    internal interface IRowWriter
    {
        /// <summary>
        /// Сохраняет строки в файл.
        /// </summary>
        /// <param name="rows">Список строк для сохранения.</param>
        /// <param name="outputPath">Путь к выходному файлу.</param>
        void Write(List<Row> rows, string outputPath);
    }
}
