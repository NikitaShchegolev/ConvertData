using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ConvertData.Application;

internal sealed class ConnectionCodeDeduplicator
{
    public int CreateDeduplicatedJson(
        string allJsonPath,
        string outputJsonPath,
        string? replacementsTxtPath = null)
    {
        if (!File.Exists(allJsonPath))
            return 0;

        var clonedRoot = JsonNode.Parse(File.ReadAllText(allJsonPath, Encoding.UTF8))!;
        if (clonedRoot is not JsonArray clonedArr)
            return 0;

        var countsByCode = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var node in clonedArr)
        {
            if (node is not JsonObject obj)
                continue;

            var code = obj["CONNECTION_CODE"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(code))
                continue;

            code = code.Trim();

            if (countsByCode.TryGetValue(code, out var c))
                countsByCode[code] = c + 1;
            else
                countsByCode[code] = 1;
        }

        var duplicateCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in countsByCode)
        {
            if (kvp.Value > 1)
                duplicateCodes.Add(kvp.Key);
        }

        var currentIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var allUsedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var changed = 0;
        var replacementsReport = replacementsTxtPath is null ? null : new StringBuilder();

        foreach (var node in clonedArr)
        {
            if (node is not JsonObject obj)
                continue;

            var code = obj["CONNECTION_CODE"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(code))
                continue;

            code = code.Trim();

            if (!duplicateCodes.Contains(code))
            {
                allUsedCodes.Add(code);
                continue;
            }

            if (!currentIndex.TryGetValue(code, out var idx))
                idx = 0;
            idx++;
            currentIndex[code] = idx;

            var newCode = $"{code}_{idx}";
            while (allUsedCodes.Contains(newCode) || countsByCode.ContainsKey(newCode))
            {
                idx++;
                currentIndex[code] = idx;
                newCode = $"{code}_{idx}";
            }

            obj["CONNECTION_CODE"] = newCode;

            replacementsReport?.AppendLine($"{code} => {newCode}");

            allUsedCodes.Add(newCode);
            changed++;
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, clonedRoot.ToJsonString(options), Encoding.UTF8);

        if (replacementsReport != null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(replacementsTxtPath!) ?? ".");
            File.WriteAllText(replacementsTxtPath!, replacementsReport.ToString(), Encoding.UTF8);
        }

        return changed;
    }
}
