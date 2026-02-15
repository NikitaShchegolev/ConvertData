using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ConvertData.Application;

internal sealed class JsonProfilePatcher
{
    public void ApplyProfilesToJson(string jsonOutDir, Dictionary<string, (double H, double B, double s, double t)> profileLookup)
    {
        foreach (var jsonPath in Directory.EnumerateFiles(jsonOutDir, "*.json", SearchOption.TopDirectoryOnly)
                     .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            PatchJsonFile(jsonPath, profileLookup);
        }
    }

    public void SelfCheckProfile(Dictionary<string, (double H, double B, double s, double t)> profileLookup)
    {
        var key = NormalizeProfileKey("10Á1");
        if (TryResolveProfile(profileLookup, key, out var g))
        {
            Console.WriteLine($"Self-check Profile=10Á1 => H={g.H}, B={g.B}, s={g.s}, t={g.t}");
            return;
        }

        Console.WriteLine("Self-check Profile=10Á1 => NOT FOUND in Profile.xls");

        var digits = new string(key.Where(char.IsDigit).ToArray());
        var sample = profileLookup.Keys
            .Where(k => !string.IsNullOrWhiteSpace(digits) && k.Contains(digits, StringComparison.OrdinalIgnoreCase))
            .Take(10)
            .ToList();

        if (sample.Count > 0)
            Console.WriteLine("Closest keys containing digits '" + digits + "': " + string.Join(", ", sample));
    }

    public static string NormalizeProfileKey(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return "";

        return new string(s
            .Trim()
            .Replace('\u00A0', ' ')
            .Where(ch => !char.IsWhiteSpace(ch))
            .ToArray());
    }

    public bool TryResolveProfile(
        Dictionary<string, (double H, double B, double s, double t)> profileLookup,
        string normalizedProfile,
        out (double H, double B, double s, double t) geometry)
    {
        if (profileLookup.TryGetValue(normalizedProfile, out geometry))
            return true;

        var digits = new string(normalizedProfile.Where(char.IsDigit).ToArray());
        if (!string.IsNullOrWhiteSpace(digits) && profileLookup.TryGetValue(digits, out geometry))
            return true;

        if (!string.IsNullOrWhiteSpace(digits))
        {
            foreach (var kv in profileLookup)
            {
                if (kv.Key.StartsWith(digits, StringComparison.OrdinalIgnoreCase))
                {
                    geometry = kv.Value;
                    return true;
                }
            }
        }

        geometry = default;
        return false;
    }

    private void PatchJsonFile(string jsonPath, Dictionary<string, (double H, double B, double s, double t)> profileLookup)
    {
        if (!TryReadJsonArray(jsonPath, out var root, out var arr))
            return;

        var patched = 0;

        foreach (var item in arr)
        {
            if (item is not JsonObject obj)
                continue;

            var key = NormalizeProfileKey(obj["Profile"]?.GetValue<string>());
            if (string.IsNullOrWhiteSpace(key) || !TryResolveProfile(profileLookup, key, out var g))
                continue;

            obj["H"] = g.H;
            obj["B"] = g.B;
            obj["s"] = g.s;
            obj["t"] = g.t;
            patched++;
        }

        if (patched == 0)
            return;

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        File.WriteAllText(jsonPath, root!.ToJsonString(options), Encoding.UTF8);
    }

    private static bool TryReadJsonArray(string jsonPath, out JsonNode? root, out JsonArray arr)
    {
        root = null;
        arr = null!;

        try
        {
            root = JsonNode.Parse(File.ReadAllText(jsonPath, Encoding.UTF8));
        }
        catch
        {
            return false;
        }

        if (root is not JsonArray a)
            return false;

        arr = a;
        return true;
    }
}
