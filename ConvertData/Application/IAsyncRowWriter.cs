using System.Collections.Generic;
using System.Threading.Tasks;
using ConvertData.Domain;

namespace ConvertData.Application
{
    /// <summary>
    /// Интерфейс для асинхронной записи списка объектов Row в выходной файл (например, JSON).
    /// </summary>
    internal interface IAsyncRowWriter
    {
        /// <summary>
        /// Асинхронно записывает список строк в файл.
        /// </summary>
        /// <param name="rows">Список строк для записи.</param>
        /// <param name="outputPath">Путь к выходному файлу.</param>
        Task WriteAsync(List<Row> rows, string outputPath);
    }
}