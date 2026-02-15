using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Объединяет все JSON-файлы из указанной папки в один файл в отдельной директории.
    /// Каждый входной файл содержит JSON-массив; результат — единый массив со всеми элементами.
    /// </summary>
    internal sealed class JsonMerger
    {
        /// <summary>
        /// Читает все `.json` из <paramref name="jsonDir"/>, объединяет в один массив
        /// и сохраняет в <paramref name="outputDir"/> под именем <paramref name="outputFileName"/>.
        /// </summary>
        public void MergeAll(string jsonDir, string outputDir, string outputFileName = "all.json")
        {
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, outputFileName);

            var merged = new JsonArray();

            var files = Directory.EnumerateFiles(jsonDir, "*.json", SearchOption.TopDirectoryOnly)
                .Where(f => !string.Equals(Path.GetFileName(f), "Profile.json", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var file in files)
            {
                try
                {
                    var text = File.ReadAllText(file, Encoding.UTF8);
                    var node = JsonNode.Parse(text);

                    if (node is JsonArray arr)
                    {
                        int count = arr.Count;
                        foreach (var item in arr.ToList())
                        {
                            arr.Remove(item);
                            merged.Add(item);
                        }
                        Console.WriteLine($"  + {Path.GetFileName(file)}: {count} records");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Merge skip: {Path.GetFileName(file)} — {ex.Message}");
                }
            }

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            File.WriteAllText(outputPath, merged.ToJsonString(options), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            Console.WriteLine($"  Merged {files.Count} files ({merged.Count} records) => {outputPath}");
        }
    }
}
