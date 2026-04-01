using System;
using System.Globalization;

namespace ConvertData.Infrastructure.Parsing;

/// <summary>
/// Парсер числовых значений, поддерживающий как русский (с запятой), так и инвариантный (с точкой) форматы.
/// </summary>
internal static class NumericParser
{
    /// <summary>
    /// Русская культура для парсинга чисел с запятой в качестве разделителя.
    /// </summary>
    private static readonly CultureInfo RuCulture = new("ru-RU");

    /// <summary>
    /// Парсит строку в значение типа double, поддерживая русский и инвариантный форматы.
    /// </summary>
    /// <param name="s">Строка для парсинга.</param>
    /// <returns>Числовое значение типа double или 0.0, если парсинг не удался.</returns>
    public static double ParseDouble(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return 0.0;

        s = s.Trim();

        if (s.Contains(',') && double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, RuCulture, out var vr))
            return vr;

        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
            return v;

        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, RuCulture, out v))
            return v;

        return 0.0;
    }

    /// <summary>
    /// Парсит строку в значение типа int, поддерживая русский и инвариантный форматы.
    /// Если прямой парсинг не удался, пытается преобразовать через double с округлением.
    /// </summary>
    /// <param name="s">Строка для парсинга.</param>
    /// <returns>Целое число или 0, если парсинг не удался.</returns>
    public static int ParseInt(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return 0;

        s = s.Trim();

        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
            return v;

        if (int.TryParse(s, NumberStyles.Integer, RuCulture, out v))
            return v;

        var d = ParseDouble(s);
        if (d != 0.0)
            return (int)Math.Round(d);

        return 0;
    }
}
