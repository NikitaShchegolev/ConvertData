using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;
using ConvertData.Domain;

namespace ConvertData.Application;

internal sealed class ProfileLookupLoader
{
    public Dictionary<string, ProfileGeometry> Load(string excelProfileDir)
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

            var dict = new Dictionary<string, ProfileGeometry>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in arr)
            {
                if (item is not JsonObject obj)
                    continue;

                var profile = obj["Profile"]?.GetValue<string>();
                var key = JsonProfilePatcher.NormalizeProfileKey(profile);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                dict[key] = new ProfileGeometry
                {
                    H = obj["H"]?.GetValue<double>() ?? 0,
                    B = obj["B"]?.GetValue<double>() ?? 0,
                    t_w = obj["t_w"]?.GetValue<double>() ?? 0,
                    t_f = obj["t_f"]?.GetValue<double>() ?? 0,
                    r1 = obj["r1"]?.GetValue<double>() ?? 0,
                    r2 = obj["r2"]?.GetValue<double>() ?? 0,
                    A = obj["A"]?.GetValue<double>() ?? 0,
                    P = obj["P"]?.GetValue<double>() ?? 0,
                    Iz = obj["Iz"]?.GetValue<double>() ?? 0,
                    Iy = obj["Iy"]?.GetValue<double>() ?? 0,
                    Ix = obj["Ix"]?.GetValue<double>() ?? 0,
                    Iv = obj["Iv"]?.GetValue<double>() ?? 0,
                    Iyz = obj["Iyz"]?.GetValue<double>() ?? 0,
                    Wz = obj["Wz"]?.GetValue<double>() ?? 0,
                    Wy = obj["Wy"]?.GetValue<double>() ?? 0,
                    Wx = obj["Wx"]?.GetValue<double>() ?? 0,
                    Wvo = obj["Wvo"]?.GetValue<double>() ?? 0,
                    Sz = obj["Sz"]?.GetValue<double>() ?? 0,
                    Sy = obj["Sy"]?.GetValue<double>() ?? 0,
                    iz = obj["iz"]?.GetValue<double>() ?? 0,
                    iy = obj["iy"]?.GetValue<double>() ?? 0,
                    xo = obj["xo"]?.GetValue<double>() ?? 0,
                    yo = obj["yo"]?.GetValue<double>() ?? 0,
                    iu = obj["iu"]?.GetValue<double>() ?? 0,
                    iv = obj["iv"]?.GetValue<double>() ?? 0,
                };
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