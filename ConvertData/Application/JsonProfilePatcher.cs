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
            Console.WriteLine($"Self-check ProfileBeam=10Á1 => H={g.H}, B={g.B}, t_w={g.t_w}, t_f={g.t_f}");
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
                beam["Beam_s"]  = bg.t_w;
                beam["Beam_t"]  = bg.t_f;
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
                column["Column_s"]  = cg.t_w;
                column["Column_t"]  = cg.t_f;
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

            // Patch Brace geometry
            var braceNode = geometryNode["Brace"];
            var braceKey = NormalizeProfileKey(braceNode?["ProfileBrace"]?.GetValue<string>());
            if (!string.IsNullOrWhiteSpace(braceKey) && TryResolveProfile(profileLookup, braceKey, out var rg) && braceNode is JsonObject brace)
            {
                brace["Brace_H"] = rg.H;
                brace["Brace_B"] = rg.B;
                brace["Brace_s"] = rg.t_w;
                brace["Brace_t"] = rg.t_f;
                brace["Brace_A"] = rg.A;
                brace["Brace_P"] = rg.P;
                brace["Brace_Iz"] = rg.Iz;
                brace["Brace_Iy"] = rg.Iy;
                brace["Brace_Ix"] = rg.Ix;
                brace["Brace_Wz"] = rg.Wz;
                brace["Brace_Wy"] = rg.Wy;
                brace["Brace_Wx"] = rg.Wx;
                brace["Brace_Sz"] = rg.Sz;
                brace["Brace_Sy"] = rg.Sy;
                brace["Brace_iz"] = rg.iz;
                brace["Brace_iy"] = rg.iy;
                brace["Brace_xo"] = rg.xo;
                brace["Brace_yo"] = rg.yo;
                itemPatched = true;
            }
            // Patch Rigel geometry
            var rigelNode = geometryNode["Rigel"];
            var rigelKey = NormalizeProfileKey(rigelNode?["ProfileRigel"]?.GetValue<string>());
            if (!string.IsNullOrWhiteSpace(rigelKey) && TryResolveProfile(profileLookup, rigelKey, out var ri) && rigelNode is JsonObject rigel)
            {
                rigel["Rigel_H"] =  ri.H;
                rigel["Rigel_B"] =  ri.B;
                rigel["Rigel_s"] =  ri.t_w;
                rigel["Rigel_t"] =  ri.t_f;
                rigel["Rigel_A"] =  ri.A;
                rigel["Rigel_P"] =  ri.P;
                rigel["Rigel_Iz"] = ri.Iz;
                rigel["Rigel_Iy"] = ri.Iy;
                rigel["Rigel_Ix"] = ri.Ix;
                rigel["Rigel_Wz"] = ri.Wz;
                rigel["Rigel_Wy"] = ri.Wy;
                rigel["Rigel_Wx"] = ri.Wx;
                rigel["Rigel_Sz"] = ri.Sz;
                rigel["Rigel_Sy"] = ri.Sy;
                rigel["Rigel_iz"] = ri.iz;
                rigel["Rigel_iy"] = ri.iy;
                rigel["Rigel_xo"] = ri.xo;
                rigel["Rigel_yo"] = ri.yo;
                itemPatched = true;
            }
            // Patch RunThrough geometry
            var runThroughNode = geometryNode["RunThrough"];
            var runThroughKey = NormalizeProfileKey(runThroughNode?["ProfileRunThrough"]?.GetValue<string>());
            if (!string.IsNullOrWhiteSpace(runThroughKey) && TryResolveProfile(profileLookup, runThroughKey, out var rt) && runThroughNode is JsonObject runThrough)
            {
                runThrough["RunThrough_H"] = rt.H;
                runThrough["RunThrough_B"] = rt.B;
                runThrough["RunThrough_s"] = rt.t_w;
                runThrough["RunThrough_t"] = rt.t_f;
                runThrough["RunThrough_A"] = rt.A;
                runThrough["RunThrough_P"] = rt.P;
                runThrough["RunThrough_Iz"] = rt.Iz;
                runThrough["RunThrough_Iy"] = rt.Iy;
                runThrough["RunThrough_Ix"] = rt.Ix;
                runThrough["RunThrough_Wz"] = rt.Wz;
                runThrough["RunThrough_Wy"] = rt.Wy;
                runThrough["RunThrough_Wx"] = rt.Wx;
                runThrough["RunThrough_Sz"] = rt.Sz;
                runThrough["RunThrough_Sy"] = rt.Sy;
                runThrough["RunThrough_iz"] = rt.iz;
                runThrough["RunThrough_iy"] = rt.iy;
                runThrough["RunThrough_xo"] = rt.xo;
                runThrough["RunThrough_yo"] = rt.yo;
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
