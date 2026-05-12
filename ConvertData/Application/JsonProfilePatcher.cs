using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ConvertData.Domain;
using ConvertData.Enums;

namespace ConvertData.Application;

internal sealed class JsonProfilePatcher
{
    private static readonly ProfileSectionType[] ProfileSections =
    [
        ProfileSectionType.Beam,
        ProfileSectionType.Column,
        ProfileSectionType.Brace,
        ProfileSectionType.Rigel,
        ProfileSectionType.RunThrougth
    ];

    private static readonly Dictionary<ProfileSectionType, ProfileSectionDefinition> SectionDefinitions = new()
    {
        [ProfileSectionType.Beam] = new("Beam", "ProfileBeam", "Beam"),
        [ProfileSectionType.Column] = new("Column", "ProfileColumn", "Column"),
        [ProfileSectionType.Brace] = new("Brace", "ProfileBrace", "Brace"),
        [ProfileSectionType.Rigel] = new("Rigel", "ProfileRigel", "Rigel"),
        [ProfileSectionType.RunThrougth] = new("RunThrougth", "ProfileRunThrough", "RunThrougth")
    };

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

        geometry = default!;
        return false;
    }

    private static void SetNumber(JsonObject target, string propertyName, double value)
    {
        target[propertyName] = value;
    }

    private static ProfileSectionDefinition? GetSectionDefinition(ProfileSectionType sectionType)
    {
        return SectionDefinitions.TryGetValue(sectionType, out var definition)
            ? definition
            : null;
    }

    private static void ApplySectionGeometry(
        JsonObject target,
        string prefix,
        ProfileGeometry geometry)
    {
        SetNumber(target, $"{prefix}_H", geometry.H);
        SetNumber(target, $"{prefix}_B", geometry.B);
        SetNumber(target, $"{prefix}_s", geometry.t_w);
        SetNumber(target, $"{prefix}_t", geometry.t_f);
        SetNumber(target, $"{prefix}_A", geometry.A);
        SetNumber(target, $"{prefix}_P", geometry.P);
        SetNumber(target, $"{prefix}_Iz", geometry.Iz);
        SetNumber(target, $"{prefix}_Iy", geometry.Iy);
        SetNumber(target, $"{prefix}_Ix", geometry.Ix);
        SetNumber(target, $"{prefix}_Wz", geometry.Wz);
        SetNumber(target, $"{prefix}_Wy", geometry.Wy);
        SetNumber(target, $"{prefix}_Wx", geometry.Wx);
        SetNumber(target, $"{prefix}_Sz", geometry.Sz);
        SetNumber(target, $"{prefix}_Sy", geometry.Sy);
        SetNumber(target, $"{prefix}_iz", geometry.iz);
        SetNumber(target, $"{prefix}_iy", geometry.iy);
        SetNumber(target, $"{prefix}_xo", geometry.xo);
        SetNumber(target, $"{prefix}_yo", geometry.yo);
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

            foreach (var sectionType in ProfileSections)
            {
                var sectionDefinition = GetSectionDefinition(sectionType);
                if (sectionDefinition is null)
                    continue;

                if (geometryNode[sectionDefinition.SectionName] is not JsonObject sectionNode)
                    continue;

                var profileKey = NormalizeProfileKey(sectionNode[sectionDefinition.ProfilePropertyName]?.GetValue<string>());
                if (string.IsNullOrWhiteSpace(profileKey) || !TryResolveProfile(profileLookup, profileKey, out var geometry))
                    continue;

                ApplySectionGeometry(sectionNode, sectionDefinition.SectionPrefix, geometry);
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

    private sealed record ProfileSectionDefinition(string SectionName, string ProfilePropertyName, string SectionPrefix);
}
