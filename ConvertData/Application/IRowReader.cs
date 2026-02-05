using System.Collections.Generic;
using ConvertData.Domain;

namespace ConvertData.Application
{
    /// <summary>
    /// Контракт чтения входного файла и получения списка строк `Row`.
    /// Реализации могут читать TSV/CSV/Excel и т.д.
    /// </summary>
    internal interface IRowReader
    {
        /// <summary>
        /// Считывает данные из файла и возвращает список строк доменной модели.
        /// </summary>
        /// <param name="path">Путь к входному файлу.</param>
        /// <returns>Список строк.</returns>
        List<Row> Read(string path);
    }
}
