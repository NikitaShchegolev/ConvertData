using System;
using System.IO;
using System.Linq;
using System.Text;
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

        private readonly ProfileLookupLoader _profileLookupLoader = new();
        private readonly JsonProfilePatcher _profilePatcher = new();

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

            var mode = RunModeParser.GetMode(args);

            if (mode == RunMode.All || mode == RunMode.CreateJson)
            {
                Console.WriteLine("=== Этап 1: Создание JSON из Excel (без профилей) ===");
                ClearJsonOut(jsonOutDir);

                foreach (var input in GetInputFiles(RunModeParser.GetInputArgsForCreateJson(args), excelDir))
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
            var allNotDuplicateJsonPath = Path.Combine(jsonAllDir, "all_NotDuplicate.json");

            Console.WriteLine();
            Console.WriteLine("=== Этап 4: Создание списков profile.txt и CONNECTION_CODE.txt ===");
            new ProfileAndConnectionCodeExporter().Export(allJsonPath, excelProfileOutDir);
            Console.WriteLine("Этап 4 завершён.");

            Console.WriteLine();
            Console.WriteLine("=== Этап 5: Создание Profile.json и CONNECTION_CODE.json ===");
            new TextListToJsonExporter().ExportProfileJson(
                Path.Combine(excelProfileOutDir, "profile.txt"),
                Path.Combine(excelProfileOutDir, "Profile.json"));
            new TextListToJsonExporter().ExportConnectionCodeJson(
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE.txt"),
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE.json"));
            Console.WriteLine("Этап 5 завершён.");

            Console.WriteLine();
            Console.WriteLine("=== Этап 6: Проверка all.json на дубликаты CONNECTION_CODE ===");
            var duplicates = new ConnectionCodeDuplicateChecker().FindDuplicates(
                allJsonPath,
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE_duplicates.txt"));
            Console.WriteLine($"Этап 6 завершён. Найдено дубликатов: {duplicates.Count}");

            Console.WriteLine();
            Console.WriteLine("=== Этап 7: Создание all_NotDuplicate.json с заменой дубликатов ===");
            var changedCodes = new ConnectionCodeDeduplicator().CreateDeduplicatedJson(
                allJsonPath,
                allNotDuplicateJsonPath,
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE_replacements.txt"));
            Console.WriteLine($"Этап 7 завершён. Заменено CONNECTION_CODE: {changedCodes}");

            Console.WriteLine();
            Console.WriteLine("=== Этап 8: Создание CONNECTION_CODE_new.json и CONNECTION_CODE_new.txt из all_NotDuplicate.json ===");
            var exporter = new ProfileAndConnectionCodeExporter();
            exporter.ExportConnectionCodesOnly(
                allNotDuplicateJsonPath,
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE_new.json"));
            var remainingDuplicates = exporter.ExportConnectionCodesTxt(
                allNotDuplicateJsonPath,
                Path.Combine(excelProfileOutDir, "CONNECTION_CODE_new.txt"));
            if (remainingDuplicates > 0)
                Console.WriteLine($"  ВНИМАНИЕ: в all_NotDuplicate.json осталось дубликатов CONNECTION_CODE: {remainingDuplicates}");
            else
                Console.WriteLine("  Проверка: дубликатов CONNECTION_CODE нет.");
            Console.WriteLine("Этап 8 завершён.");
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
            var profileLookup = _profileLookupLoader.Load(excelProfileDir);
            if (profileLookup.Count == 0)
            {
                Console.WriteLine("Profile lookup is empty: EXCEL_Profile/Profile.xls was not parsed.");
                return;
            }

            _profilePatcher.SelfCheckProfile(profileLookup);
            _profilePatcher.ApplyProfilesToJson(jsonOutDir, profileLookup);
        }
    }
}
