using System;
using System.Linq;

namespace ConvertData.Application;

internal static class RunModeParser
{
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

    public static string[] GetInputArgsForCreateJson(string[] args)
    {
        if (args.Length == 0)
            return args;

        if (string.Equals(args[0], "1", StringComparison.OrdinalIgnoreCase))
            return args.Skip(1).ToArray();

        return args;
    }
}
