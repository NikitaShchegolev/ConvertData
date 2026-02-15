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

        var maxByPrefix = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
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

            if (TryParsePrefixedNumber(code, out var prefix, out var number))
            {
                if (!maxByPrefix.TryGetValue(prefix, out var currentMax) || number > currentMax)
                    maxByPrefix[prefix] = number;
            }
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

            if (!countsByCode.TryGetValue(code, out var total) || total <= 1)
            {
                seen.Add(code);
                continue;
            }

            if (seen.Add(code))
                continue;

            if (!TryParsePrefixedNumber(code, out var prefix, out _))
                continue;

            maxByPrefix.TryGetValue(prefix, out var max);
            var next = max + 1;
            string newCode;
            do
            {
                newCode = prefix + "-" + next;
                next++;
            }
            while (seen.Contains(newCode) || countsByCode.ContainsKey(newCode));

            obj["CONNECTION_CODE"] = newCode;

            replacementsReport?.AppendLine($"{code} => {newCode}");

            seen.Add(newCode);
            changed++;

            maxByPrefix[prefix] = int.Parse(newCode.AsSpan(prefix.Length + 1));
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

    private static bool TryParsePrefixedNumber(string code, out string prefix, out int number)
    {
        prefix = "";
        number = 0;

        var dash = code.IndexOf('-');
        if (dash <= 0 || dash == code.Length - 1)
            return false;

        prefix = code[..dash].Trim();
        var numPart = code[(dash + 1)..].Trim();
        return !string.IsNullOrWhiteSpace(prefix) && int.TryParse(numPart, out number) && number >= 0;
    }
}
