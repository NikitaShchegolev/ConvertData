using System;
using System.Linq;

namespace ConvertData.Application;

/// <summary>
/// Парсер аргументов командной строки для определения режима выполнения приложения.
/// </summary>
internal static class RunModeParser
{
    /// <summary>
    /// Определяет режим выполнения на основе аргументов командной строки.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Режим выполнения: All (по умолчанию), CreateJson ("1") или ApplyProfiles ("2").</returns>
    public static RunMode GetMode(string[] args)
    {
        if (args.Length == 0)
            return RunMode.All;

        if (args.Length >= 1 && string.Equals(args[0], "1", StringComparison.OrdinalIgnoreCase))
            return RunMode.CreateJson;

        if (args.Length >= 1 && string.Equals(args[0], "2", StringComparison.OrdinalIgnoreCase))
            return RunMode.ApplyProfiles;

        return RunMode.All;
    }

    /// <summary>
    /// Извлекает аргументы для режима CreateJson, пропуская первый аргумент-флаг режима.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Массив аргументов без флага режима.</returns>
    public static string[] GetInputArgsForCreateJson(string[] args)
    {
        if (args.Length == 0)
            return args;

        if (string.Equals(args[0], "1", StringComparison.OrdinalIgnoreCase))
            return args.Skip(1).ToArray();

        return args;
    }

    /// <summary>
    /// Извлекает значение параметра --profile-column из аргументов командной строки.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Имя колонки профиля или null, если параметр не указан.</returns>
    public static string? GetProfileColumn(string[] args)
    {
        const string prefix = "--profile-column=";
        foreach (var arg in args)
        {
            if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return arg.Substring(prefix.Length);
        }
        return null;
    }
}
