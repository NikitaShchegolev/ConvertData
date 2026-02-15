using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ConvertData.Entitys;

namespace ConvertData.Application;

internal sealed class NameConnectionsExporter
{
    public void Export(string jsonAllFilePath, string outputJsonPath)
    {
        if (!File.Exists(jsonAllFilePath))
            return;

        JsonNode? root;
        try
        {
            root = JsonNode.Parse(File.ReadAllText(jsonAllFilePath, Encoding.UTF8));
        }
        catch
        {
            return;
        }

        if (root is not JsonArray arr)
            return;

        var names = new SortedSet<string>(StringComparer.Ordinal);

        foreach (var node in arr)
        {
            if (node is not JsonObject obj)
                continue;

            var name = obj["Name"]?.GetValue<string>();
            names.Add(string.IsNullOrWhiteSpace(name) ? "Empty" : name.Trim());
        }

        var items = names
            .Select(n => new NameItem
            {
                NAME_GUID = Guid.NewGuid(),
                Name = n
            })
            .ToArray();

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, JsonSerializer.Serialize(items, options), Encoding.UTF8);
    }
}
