using System;
using System.Linq;
using System.Collections.Generic;

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
    /// Определяет, какие блоки выполнять на основе аргументов командной строки.
    /// Если указан параметр --blocks, используется он. Иначе используется режим GetMode для обратной совместимости.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Флаги блоков для выполнения.</returns>
    public static Block GetBlocks(string[] args)
    {
        // Парсим параметр --blocks
        var blocksParam = GetParameterValue(args, "--blocks");
        if (!string.IsNullOrEmpty(blocksParam))
        {
            return ParseBlocks(blocksParam);
        }

        // Парсим параметр --skip-blocks
        var skipParam = GetParameterValue(args, "--skip-blocks");
        if (!string.IsNullOrEmpty(skipParam))
        {
            var allBlocks = Block.All;
            var skipBlocks = ParseBlocks(skipParam);
            return allBlocks & ~skipBlocks;
        }

        // Если есть аргументы, не являющиеся параметрами (не начинаются с --), попробуем интерпретировать как номера блоков
        var nonParamArgs = args.Where(a => !a.StartsWith("--")).ToArray();
        if (nonParamArgs.Length > 0)
        {
            // Объединяем аргументы через запятую (например, "12" или "7,8")
            var combined = string.Join(",", nonParamArgs);
            var blocks = ParseBlocks(combined);
            if (blocks != Block.None)
                return blocks;
        }

        // Обратная совместимость: используем старый режим
        var mode = GetMode(args);
        return mode switch
        {
            RunMode.CreateJson => Block.Conversion, // Только конвертация (этапы 1-2)
            RunMode.ApplyProfiles => Block.Conversion | Block.Processing, // Конвертация + обработка (но без анкеров)
            _ => Block.All // Все блоки
        };
    }

    /// <summary>
    /// Парсит строку с названиями блоков (например, "Conversion,Processing" или "1,2").
    /// </summary>
    private static Block ParseBlocks(string blocksStr)
    {
        if (string.IsNullOrWhiteSpace(blocksStr))
            return Block.None;

        var blocks = Block.None;
        var parts = blocksStr.Split(',', ';', '|', ' ')
            .Select(s => s.Trim())
            .Where(s => !string.IsNullOrEmpty(s));

        foreach (var part in parts)
        {
            // Попробуем распарсить как число (номер блока: 1=Conversion, 2=Processing, 3=Anchors)
            if (int.TryParse(part, out int blockNumber))
            {
                var block = BlockFromNumber(blockNumber);
                if (block != Block.None)
                    blocks |= block;
                continue;
            }

            // Попробуем распарсить как имя блока
            if (Enum.TryParse<Block>(part, true, out var namedBlock))
            {
                blocks |= namedBlock;
                continue;
            }
        }

        return blocks;
    }

    /// <summary>
    /// Преобразует номер блока (1-13) в значение Block.
    /// </summary>
    private static Block BlockFromNumber(int number)
    {
        return number switch
        {
            1 => Block.CreateJson,
            2 => Block.ApplyProfiles,
            3 => Block.MergeAndEnrich,
            4 => Block.ExportProfiles,
            5 => Block.Deduplication,
            6 => Block.CopyToData,
            7 => Block.AnchorExport,
            8 => Block.SteelExport,
            9 => Block.Conversion,
            10 => Block.Processing,
            11 => Block.Anchors,
            12 => Block.Bolts,
            13 => Block.All,
            _ => Block.None
        };
    }

    /// <summary>
    /// Извлекает значение параметра из аргументов командной строки.
    /// </summary>
    private static string? GetParameterValue(string[] args, string paramName)
    {
        foreach (var arg in args)
        {
            if (arg.StartsWith(paramName + "=", StringComparison.OrdinalIgnoreCase))
                return arg.Substring(paramName.Length + 1);
        }
        return null;
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
