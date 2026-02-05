using System;
using System.IO;
using System.Linq;
using ConvertData.Application;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Инфраструктурный компонент для определения путей.
    /// Используется для нахождения папки проекта и для проверки расширений входных файлов.
    /// </summary>
    internal sealed class PathResolver : IPathResolver
    {
        /// <summary>
        /// Поднимается по дереву директорий вверх от `startDir` и ищет каталог, где лежит файл `.csproj`.
        /// </summary>
        /// <param name="startDir">Стартовая директория (обычно `bin/...`).</param>
        /// <returns>Путь к директории проекта или `null`.</returns>
        public string? GetProjectDir(string startDir)
        {
            try
            {
                var dir = new DirectoryInfo(startDir);
                for (int i = 0; i < 10 && dir != null; i++, dir = dir.Parent)
                {
                    if (dir.EnumerateFiles("*.csproj", SearchOption.TopDirectoryOnly).Any())
                        return dir.FullName;
                }
            }
            catch
            {
                // Игнорируем ошибки доступа/IO и возвращаем null.
            }

            return null;
        }

        /// <summary>
        /// Проверяет, является ли файл поддерживаемым входным "excel" файлом по расширению.
        /// В текущей логике обрабатываются только `.xls`.
        /// </summary>
        /// <param name="path">Путь к файлу.</param>
        /// <returns>`true`, если расширение поддерживается.</returns>
        public static bool HasExcelExtension(string path)
        {
            var ext = Path.GetExtension(path);
            return string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase);
        }
    }
}
