using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ConvertData.Domain;

namespace ConvertData.Application;

internal sealed class JsonProfilePatcher
{
    public void ApplyProfilesToJson(string jsonOutDir, Dictionary<string, ProfileGeometry> profileLookup)
    {
        foreach (var jsonPath in Directory.EnumerateFiles(jsonOutDir, "*.json", SearchOption.TopDirectoryOnly)
                     .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            PatchJsonFile(jsonPath, profileLookup);
        }
    }

    public void SelfCheckProfile(Dictionary<string, ProfileGeometry> profileLookup)
    {
        var key = NormalizeProfileKey("10Á1");
        if (TryResolveProfile(profileLookup, key, out var g))
        {
            Console.WriteLine($"Self-check ProfileBeam=10Á1 => H={g.H}, B={g.B}, s={g.s}, t={g.t}");
            return;
        }

        Console.WriteLine("Self-check ProfileBeam=10Á1 => NOT FOUND in ProfileBeam.xls");

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
        Dictionary<string, ProfileGeometry> profileLookup,
        string normalizedProfile,
        out ProfileGeometry geometry)
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

        geometry = default!;
        return false;
    }

    private void PatchJsonFile(string jsonPath, Dictionary<string, ProfileGeometry> profileLookup)
    {
        if (!TryReadJsonArray(jsonPath, out var root, out var arr))
            return;

        var patched = 0;

        foreach (var item in arr)
        {
            if (item is not JsonObject obj)
                continue;

            var geometryNode = obj["Geometry"] as JsonObject;
            if (geometryNode is null)
                continue;

            bool itemPatched = false;

            // Patch Beam geometry
            var beamNode = geometryNode["Beam"];
            var beamKey = NormalizeProfileKey(beamNode?["ProfileBeam"]?.GetValue<string>());
            if (!string.IsNullOrWhiteSpace(beamKey) && TryResolveProfile(profileLookup, beamKey, out var bg) && beamNode is JsonObject beam)
            {
                beam["Beam_H"]  = bg.H;
                beam["Beam_B"]  = bg.B;
                beam["Beam_s"]  = bg.s;
                beam["Beam_t"]  = bg.t;
                beam["Beam_A"]  = bg.A;
                beam["Beam_P"]  = bg.P;
                beam["Beam_Iz"] = bg.Iz;
                beam["Beam_Iy"] = bg.Iy;
                beam["Beam_Ix"] = bg.Ix;
                beam["Beam_Wz"] = bg.Wz;
                beam["Beam_Wy"] = bg.Wy;
                beam["Beam_Wx"] = bg.Wx;
                beam["Beam_Sz"] = bg.Sz;
                beam["Beam_Sy"] = bg.Sy;
                beam["Beam_iz"] = bg.iz;
                beam["Beam_iy"] = bg.iy;
                beam["Beam_xo"] = bg.xo;
                beam["Beam_yo"] = bg.yo;
                itemPatched = true;
            }

            // Patch Column geometry
            var columnNode = geometryNode["Column"];
            var columnKey = NormalizeProfileKey(columnNode?["ProfileColumn"]?.GetValue<string>());
            if (!string.IsNullOrWhiteSpace(columnKey) && TryResolveProfile(profileLookup, columnKey, out var cg) && columnNode is JsonObject column)
            {
                column["Column_H"]  = cg.H;
                column["Column_B"]  = cg.B;
                column["Column_s"]  = cg.s;
                column["Column_t"]  = cg.t;
                column["Column_A"]  = cg.A;
                column["Column_P"]  = cg.P;
                column["Column_Iz"] = cg.Iz;
                column["Column_Iy"] = cg.Iy;
                column["Column_Ix"] = cg.Ix;
                column["Column_Wz"] = cg.Wz;
                column["Column_Wy"] = cg.Wy;
                column["Column_Wx"] = cg.Wx;
                column["Column_Sz"] = cg.Sz;
                column["Column_Sy"] = cg.Sy;
                column["Column_iz"] = cg.iz;
                column["Column_iy"] = cg.iy;
                column["Column_xo"] = cg.xo;
                column["Column_yo"] = cg.yo;
                itemPatched = true;
            }

            if (itemPatched) patched++;
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
