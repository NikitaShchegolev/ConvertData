using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ConvertData.Domain;
using ConvertData.Infrastructure;

namespace ConvertData.Application
{
    internal sealed class ConvertApp
    {
        private readonly IRowWriter _writer = new JsonRowWriter();
        private readonly IRowReaderFactory _readerFactory = new RowReaderFactory();
        private readonly IPathResolver _pathResolver = new PathResolver();
        private readonly ILicenseConfigurator _licenseConfigurator = new EpplusLicenseConfigurator();

        public void Run(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _licenseConfigurator.Configure();

            var projectDir = _pathResolver.GetProjectDir(AppDomain.CurrentDomain.BaseDirectory) ?? AppDomain.CurrentDomain.BaseDirectory;
            var excelDir = Path.Combine(projectDir, "EXCEL");
            var excelProfileDir = Path.Combine(projectDir, "EXCEL_Profile");
            var jsonOutDir = Path.Combine(projectDir, "JSON_OUT");
            var jsonAllDir = Path.Combine(projectDir, "JSON_All");
            var excelProfileOutDir = Path.Combine(projectDir, "EXCEL_Profile_OUT");
            Directory.CreateDirectory(jsonOutDir);

            var mode = GetMode(args);

            if (mode == RunMode.All || mode == RunMode.CreateJson)
            {
                Console.WriteLine("=== Этап 1: Создание JSON из Excel (без профилей) ===");
                ClearJsonOut(jsonOutDir);

                foreach (var input in GetInputFiles(GetInputArgsForCreateJson(args), excelDir))
                    ConvertOne(input, jsonOutDir);

                Console.WriteLine("Этап 1 завершён.");                
            }

            if (mode == RunMode.All || mode == RunMode.ApplyProfiles)
            {
                Console.WriteLine();
                Console.WriteLine("=== Этап 2: Применение справочника профилей (H, B, s, t) ===");
                ApplyProfilesToJson(jsonOutDir, excelProfileDir);
                Console.WriteLine("Этап 2 завершён.");
            }

            Console.WriteLine();
            Console.WriteLine("=== Этап 3: Объединение всех JSON в один файл ===");
            new JsonMerger().MergeAll(jsonOutDir, jsonAllDir);
            Console.WriteLine("Этап 3 завершён.");

            var allJsonPath = Path.Combine(jsonAllDir, "all.json");
            new ProfileAndConnectionCodeExporter().Export(allJsonPath, excelProfileOutDir);
        }

        private enum RunMode
        {
            All,
            CreateJson,
            ApplyProfiles
        }

        private static RunMode GetMode(string[] args)
        {
            if (args.Length == 0)
                return RunMode.All;

            if (args.Length >= 1 && string.Equals(args[0], "1", StringComparison.OrdinalIgnoreCase))
                return RunMode.CreateJson;

            if (args.Length >= 1 && string.Equals(args[0], "2", StringComparison.OrdinalIgnoreCase))
                return RunMode.ApplyProfiles;

            return RunMode.All;
        }

        private static string[] GetInputArgsForCreateJson(string[] args)
        {
            if (args.Length == 0)
                return args;

            if (string.Equals(args[0], "1", StringComparison.OrdinalIgnoreCase))
                return args.Skip(1).ToArray();

            return args;
        }

        private static void ClearJsonOut(string jsonOutDir)
        {
            foreach (var f in Directory.EnumerateFiles(jsonOutDir, "*.json", SearchOption.TopDirectoryOnly))
            {
                try { File.Delete(f); }
                catch { }
            }
        }

        private static IEnumerable<string> GetInputFiles(string[] args, string excelDir)
        {
            if (args.Length > 0)
            {
                return args.Where(p => !string.Equals(Path.GetFileName(p), "Profile.xls", StringComparison.OrdinalIgnoreCase));
            }

            if (!Directory.Exists(excelDir))
                throw new DirectoryNotFoundException("Input folder not found: " + excelDir);

            return Directory.EnumerateFiles(excelDir)
                .Where(PathResolver.HasExcelExtension)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
        }

        private void ConvertOne(string inputPath, string jsonOutDir)
        {
            if (!File.Exists(inputPath))
                return;

            var rows = _readerFactory.Create(inputPath).Read(inputPath);

            var outPath = Path.Combine(jsonOutDir, Path.GetFileNameWithoutExtension(inputPath) + ".json");
            _writer.Write(rows, outPath);
            Console.WriteLine("Written: " + outPath);
        }

        private void ApplyProfilesToJson(string jsonOutDir, string excelProfileDir)
        {
            var profileLookup = BuildProfileLookup(excelProfileDir);
            if (profileLookup.Count == 0)
            {
                Console.WriteLine("Profile lookup is empty: EXCEL_Profile/Profile.xls was not parsed.");
                return;
            }

            SelfCheckProfile(profileLookup);

            foreach (var jsonPath in Directory.EnumerateFiles(jsonOutDir, "*.json", SearchOption.TopDirectoryOnly)
                         .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                PatchJsonFile(jsonPath, profileLookup);
            }
        }

        private static void SelfCheckProfile(Dictionary<string, (double H, double B, double s, double t)> profileLookup)
        {
            var key = NormalizeProfileKey("10Б1");
            if (TryResolveProfile(profileLookup, key, out var g))
            {
                Console.WriteLine($"Self-check Profile=10Б1 => H={g.H}, B={g.B}, s={g.s}, t={g.t}");
                return;
            }

            Console.WriteLine("Self-check Profile=10Б1 => NOT FOUND in Profile.xls");

            var digits = new string(key.Where(char.IsDigit).ToArray());
            var sample = profileLookup.Keys
                .Where(k => !string.IsNullOrWhiteSpace(digits) && k.Contains(digits, StringComparison.OrdinalIgnoreCase))
                .Take(10)
                .ToList();

            if (sample.Count > 0)
                Console.WriteLine("Closest keys containing digits '" + digits + "': " + string.Join(", ", sample));
        }

        private static void PatchJsonFile(string jsonPath, Dictionary<string, (double H, double B, double s, double t)> profileLookup)
        {
            if (!TryReadJsonArray(jsonPath, out var root, out var arr))
                return;

            var patched = 0;

            foreach (var item in arr)
            {
                if (item is not JsonObject obj)
                    continue;

                var key = NormalizeProfileKey(obj["Profile"]?.GetValue<string>());
                if (string.IsNullOrWhiteSpace(key) || !TryResolveProfile(profileLookup, key, out var g))
                    continue;

                obj["H"] = g.H;
                obj["B"] = g.B;
                obj["s"] = g.s;
                obj["t"] = g.t;
                patched++;
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

        private Dictionary<string, (double H, double B, double s, double t)> BuildProfileLookup(string excelProfileDir)
        {
            var profilePath = Path.Combine(excelProfileDir, "Profile.json");
            if (!File.Exists(profilePath))
                return new(StringComparer.OrdinalIgnoreCase);

            try
            {
                var json = File.ReadAllText(profilePath, Encoding.UTF8);
                var arr = JsonNode.Parse(json) as JsonArray;
                if (arr is null)
                    return new(StringComparer.OrdinalIgnoreCase);

                var dict = new Dictionary<string, (double H, double B, double s, double t)>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in arr)
                {
                    if (item is not JsonObject obj)
                        continue;

                    var profile = obj["Profile"]?.GetValue<string>();
                    var key = NormalizeProfileKey(profile);
                    if (string.IsNullOrWhiteSpace(key))
                        continue;

                    double h = obj["H"]?.GetValue<double>() ?? 0;
                    double b = obj["B"]?.GetValue<double>() ?? 0;
                    double s = obj["s"]?.GetValue<double>() ?? 0;
                    double t = obj["t"]?.GetValue<double>() ?? 0;

                    dict[key] = (h, b, s, t);
                }

                Console.WriteLine($"  Loaded profiles: {dict.Count} from {profilePath}");
                return dict;
            }
            catch (Exception ex)
            {
                Console.WriteLine("  Failed to read profile json: " + profilePath);
                Console.WriteLine(ex);
                return new(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static string NormalizeProfileKey(string? s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "";

            return new string(s
                .Trim()
                .Replace('\u00A0', ' ')
                .Where(ch => !char.IsWhiteSpace(ch))
                .ToArray());
        }

        private static bool TryResolveProfile(
            Dictionary<string, (double H, double B, double s, double t)> profileLookup,
            string normalizedProfile,
            out (double H, double B, double s, double t) geometry)
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

            geometry = default;
            return false;
        }
    }
}
