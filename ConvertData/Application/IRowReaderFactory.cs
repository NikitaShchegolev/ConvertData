namespace ConvertData.Application
{
    /// <summary>
    /// Фабрика, выбирающая подходящую реализацию `IRowReader` для конкретного файла.
    /// </summary>
    internal interface IRowReaderFactory
    {
        /// <summary>
        /// Создаёт reader на основании пути/расширения/имени файла.
        /// </summary>
        /// <param name="path">Путь к входному файлу.</param>
        /// <returns>Reader, который сможет прочитать данный файл.</returns>
        IRowReader Create(string path);
    }
}
