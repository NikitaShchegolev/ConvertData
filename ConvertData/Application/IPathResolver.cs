namespace ConvertData.Application
{
    /// <summary>
    /// Контракт для поиска/разрешения путей, необходимых приложению.
    /// </summary>
    internal interface IPathResolver
    {
        /// <summary>
        /// Пытается найти папку проекта (где лежит `.csproj`), поднимаясь вверх от стартовой директории.
        /// </summary>
        /// <param name="startDir">Стартовая директория (обычно `AppDomain.CurrentDomain.BaseDirectory`).</param>
        /// <returns>Путь к папке проекта или `null`, если не удалось определить.</returns>
        string? GetProjectDir(string startDir);
    }
}
