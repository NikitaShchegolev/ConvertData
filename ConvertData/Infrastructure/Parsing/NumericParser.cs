using System;
using System.Globalization;

namespace ConvertData.Infrastructure.Parsing;

internal static class NumericParser
{
    private static readonly CultureInfo RuCulture = new("ru-RU");

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
