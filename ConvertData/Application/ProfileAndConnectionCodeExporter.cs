using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ConvertData.Entitys;

namespace ConvertData.Application
{
    internal sealed class ProfileAndConnectionCodeExporter
    {
        public void Export(string jsonAllFilePath, string outDir)
        {
            if (!File.Exists(jsonAllFilePath))
                return;

            Directory.CreateDirectory(outDir);

            if (!TryReadJsonArray(jsonAllFilePath, out var arr))
                return;

            var profiles = new SortedSet<string>(StringComparer.Ordinal);
            var connectionCodes = new SortedSet<string>(StringComparer.Ordinal);

            foreach (var node in arr)
            {
                if (node is not JsonObject obj)
                    continue;

                var profile = obj["Profile"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(profile))
                    profiles.Add(profile.Trim());

                var connectionCode = obj["CONNECTION_CODE"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(connectionCode))
                    connectionCodes.Add(connectionCode.Trim());
            }

            File.WriteAllLines(Path.Combine(outDir, "profile.txt"), profiles, Encoding.UTF8);
            File.WriteAllLines(Path.Combine(outDir, "CONNECTION_CODE.txt"), connectionCodes, Encoding.UTF8);
        }

        public void ExportConnectionCodesOnly(string jsonAllFilePath, string outputJsonPath)
        {
            if (!File.Exists(jsonAllFilePath))
                return;

            if (!TryReadJsonArray(jsonAllFilePath, out var arr))
                return;

            var connectionCodes = new SortedSet<string>(StringComparer.Ordinal);

            foreach (var node in arr)
            {
                if (node is not JsonObject obj)
                    continue;

                var connectionCode = obj["CONNECTION_CODE"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(connectionCode))
                    connectionCodes.Add(connectionCode.Trim());
            }

            var items = connectionCodes
                .Select(code => new ConnectionCodeItem
                {
                    CONNECTION_GUID = Guid.NewGuid(),
                    CONNECTION_CODE = code
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

        public int ExportConnectionCodesTxt(string jsonAllFilePath, string outputTxtPath)
        {
            if (!File.Exists(jsonAllFilePath))
                return 0;

            if (!TryReadJsonArray(jsonAllFilePath, out var arr))
                return 0;

            var allCodes = new List<string>();

            foreach (var node in arr)
            {
                if (node is not JsonObject obj)
                    continue;

                var connectionCode = obj["CONNECTION_CODE"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(connectionCode))
                    allCodes.Add(connectionCode.Trim());
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var duplicates = new List<string>();
            foreach (var code in allCodes)
            {
                if (!seen.Add(code))
                    duplicates.Add(code);
            }

            var uniqueSorted = new SortedSet<string>(allCodes, StringComparer.Ordinal);

            Directory.CreateDirectory(Path.GetDirectoryName(outputTxtPath) ?? ".");
            File.WriteAllLines(outputTxtPath, uniqueSorted, Encoding.UTF8);

            return duplicates.Count;
        }

        private static bool TryReadJsonArray(string jsonPath, out JsonArray arr)
        {
            arr = null!;

            JsonNode? root;
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
}
