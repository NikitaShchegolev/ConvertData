using System;
using System.Collections.Generic;

namespace ConvertData.Infrastructure.Parsing;

internal static class HeaderUtils
{
    public static string NormalizeHeader(string h)
    {
        if (string.IsNullOrWhiteSpace(h))
            return string.Empty;

        return h.Trim();
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
