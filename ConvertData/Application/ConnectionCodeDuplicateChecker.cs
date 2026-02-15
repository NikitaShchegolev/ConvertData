using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;

namespace ConvertData.Application;

internal sealed class ConnectionCodeDuplicateChecker
{
    public List<string> FindDuplicates(string allJsonPath, string? duplicatesTxtPath = null)
    {
        var duplicates = new List<string>();

        if (!File.Exists(allJsonPath))
            return duplicates;

        var root = JsonNode.Parse(File.ReadAllText(allJsonPath, Encoding.UTF8));
        if (root is not JsonArray arr)
            return duplicates;

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var node in arr)
        {
            if (node is not JsonObject obj)
                continue;

            var code = obj["CONNECTION_CODE"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(code))
                continue;

            code = code.Trim();

            if (!seen.Add(code))
                duplicates.Add(code);
        }

        if (duplicatesTxtPath != null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(duplicatesTxtPath) ?? ".");
            File.WriteAllLines(duplicatesTxtPath, duplicates, Encoding.UTF8);
        }

        return duplicates;
    }
}
