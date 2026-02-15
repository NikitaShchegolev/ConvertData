using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using ConvertData.Entitys;

namespace ConvertData.Application
{
    internal sealed class TextListToJsonExporter
    {
        public void ExportProfileJson(string inputProfileTxtPath, string outputProfileJsonPath)
        {
            var values = ReadDistinctNonEmptyLines(inputProfileTxtPath);

            var arr = values
                .Select(v => new ProfileItem
                {
                    CONNECTION_GUID = Guid.NewGuid(),
                    Profile = v
                })
                .ToArray();

            WriteJson(outputProfileJsonPath, arr);
        }

        public void ExportConnectionCodeJson(string inputConnectionCodeTxtPath, string outputConnectionCodeJsonPath)
        {
            var values = ReadDistinctNonEmptyLines(inputConnectionCodeTxtPath);

            var arr = values
                .Select(v => new ConnectionCodeItem
                {
                    CONNECTION_GUID = Guid.NewGuid(),
                    CONNECTION_CODE = v
                })
                .ToArray();

            WriteJson(outputConnectionCodeJsonPath, arr);
        }

        private static List<string> ReadDistinctNonEmptyLines(string path)
        {
            if (!File.Exists(path))
                return new();

            return File.ReadAllLines(path, Encoding.UTF8)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToList();
        }

        private static void WriteJson<T>(string path, T value)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            File.WriteAllText(path, JsonSerializer.Serialize(value, options), Encoding.UTF8);
        }
    }
}
