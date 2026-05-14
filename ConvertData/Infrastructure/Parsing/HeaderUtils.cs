using System;
using System.Collections.Generic;

namespace ConvertData.Infrastructure.Parsing;
/// <summary>
/// нужен для нормализации и поиска заголовков Excel
/// •	приводит заголовки к единому виду
/// •	убирает мусорные/невидимые символы
/// •	обрезает отличия, из-за которых одинаковые заголовки могут не совпасть строка в строку
/// </summary>
internal static class HeaderUtils
{
    public static string NormalizeHeader(string h)
    {
        if (string.IsNullOrWhiteSpace(h))
            return string.Empty;

        var sb = new System.Text.StringBuilder(h.Length);
        foreach (var ch in h)
        {
            if (ch == '\u00A0' || ch == '\uFEFF' || ch == '\u200B' || ch == '\u200C' || ch == '\u200D')
                continue;
            sb.Append(ch);
        }

        return sb.ToString().Trim();
    }

    public static int IndexOfHeader(List<string> header, string name)
    {
        for (int i = 0; i < header.Count; i++)
        {
            if (string.Equals(header[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        }
        return -1;
    }

    public static int IndexOfHeaderAny(List<string> header, IEnumerable<string> names)
    {
        foreach (var n in names)
        {
            int idx = IndexOfHeader(header, n);
            if (idx >= 0)
                return idx;
        }
        return -1;
    }
}
