using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;

namespace ConvertData.Application;

internal sealed class ProfileLookupLoader
{
    public Dictionary<string, (double H, double B, double s, double t)> Load(string excelProfileDir)
    {
        var profilePath = Path.Combine(excelProfileDir, "Profile.json");
        if (!File.Exists(profilePath))
            return new(StringComparer.OrdinalIgnoreCase);

        try
        {
            var json = File.ReadAllText(profilePath, Encoding.UTF8);
            var arr = JsonNode.Parse(json) as JsonArray;
            if (arr is null)
                return new(StringComparer.OrdinalIgnoreCase);

            var dict = new Dictionary<string, (double H, double B, double s, double t)>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in arr)
            {
                if (item is not JsonObject obj)
                    continue;

                var profile = obj["Profile"]?.GetValue<string>();
                var key = JsonProfilePatcher.NormalizeProfileKey(profile);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                double h = obj["H"]?.GetValue<double>() ?? 0;
                double b = obj["B"]?.GetValue<double>() ?? 0;
                double s = obj["s"]?.GetValue<double>() ?? 0;
                double t = obj["t"]?.GetValue<double>() ?? 0;

                dict[key] = (h, b, s, t);
            }

            Console.WriteLine($"  Loaded profiles: {dict.Count} from {profilePath}");
            return dict;
        }
        catch (Exception ex)
        {
            Console.WriteLine("  Failed to read profile json: " + profilePath);
            Console.WriteLine(ex);
            return new(StringComparer.OrdinalIgnoreCase);
        }
    }
}
