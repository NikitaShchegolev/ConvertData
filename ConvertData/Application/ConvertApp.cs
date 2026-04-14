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
        private readonly ProfileExcelToJsonExporter _profileExcelExporter = new();
        private readonly SteelExcelToJsonExporter _anchorExcelExporter = new();

        public void Run(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _licenseConfigurator.Configure();

            var projectDir = _pathResolver.GetProjectDir(AppDomain.CurrentDomain.BaseDirectory) ?? AppDomain.CurrentDomain.BaseDirectory;
            var excelDir = Path.Combine(projectDir, "EXCEL");
            var excelProfileDir = Path.Combine(projectDir, "EXCEL_Profile");
            // Allow overriding profile column header via command-line
            var profileColumn = RunModeParser.GetProfileColumn(args);
            if (!string.IsNullOrWhiteSpace(profileColumn))
                ExcelHeaderResolver.ProfileColumnOverride = profileColumn.Trim();
            var jsonOutDir = Path.Combine(projectDir, "JSON_OUT");
            var jsonAllDir = Path.Combine(projectDir, "JSON_All");
            var excelProfileOutDir = Path.Combine(projectDir, "EXCEL_Profile_OUT");
            var excelAnchorDir = Path.Combine(projectDir, "EXCEL_Anchor");
            var excelAnchorOutDir = Path.Combine(projectDir, "EXCEL_Anchor_OUT");
            var exceSteelDir = Path.Combine(projectDir, "EXCEL_MARK_STEEL");
            var exceSteelDirOut = Path.Combine(projectDir, "EXCEL_MARK_STEEL_OUT");
            Directory.CreateDirectory(jsonOutDir);

            var stages = RunModeParser.GetStages(args);
            Console.WriteLine($"Выполняемые этапы: {stages}");

            // Этап 1: Создание JSON из Excel (без профилей)
            if (stages.HasFlag(Stage.CreateJsonFromExcel))
            {
                Console.WriteLine("=== Этап 1: Создание JSON из Excel (без профилей) ===");
                ClearJsonOut(jsonOutDir);

                foreach (var input in GetInputFiles(RunModeParser.GetInputArgsForCreateJson(args), excelDir))
                    ConvertOne(input, jsonOutDir);

                Console.WriteLine("Этап 1 завершён.");
            }
            else if (stages.HasFlag(Stage.ApplyProfiles) || stages.HasFlag(Stage.MergeJson))
            {
                // Если пропущен этап 1, но нужны последующие, проверяем наличие JSON файлов
                if (!Directory.Exists(jsonOutDir) || !Directory.EnumerateFiles(jsonOutDir, "*.json").Any())
                {
                    Console.WriteLine("ВНИМАНИЕ: Этап 1 пропущен, но в папке JSON_OUT нет JSON файлов.");
                    Console.WriteLine("Для выполнения последующих этапов нужны JSON файлы.");
                }
            }

            // Этап 2: Применение справочника профилей
            if (stages.HasFlag(Stage.ApplyProfiles))
            {
                Console.WriteLine();
                Console.WriteLine("=== Этап 2: Применение справочника профилей (Beam_H, Beam_B, Beam_s, Beam_t) ===");
                ApplyProfilesToJson(jsonOutDir, excelProfileDir);
                Console.WriteLine("Этап 2 завершён.");

                Console.WriteLine();
                Console.WriteLine("=== Шаг 1.5: Экспорт профилей из Excel → Profile.json ===");
                _profileExcelExporter.Export(
                    excelProfileDir,
                    Path.Combine(excelProfileOutDir, "Profile.json"));
                Console.WriteLine("Шаг 1.5 завершён.");
            }

            // Этап 3: Объединение всех JSON в один файл
            if (stages.HasFlag(Stage.MergeJson))
            {
                Console.WriteLine();
                Console.WriteLine("=== Этап 3: Объединение всех JSON в один файл ===");
                var merged = new JsonMerger().MergeAll(jsonOutDir);
                Console.WriteLine("Этап 3 завершён.");

                // Этап 4: Обогащение неполных записей
                if (stages.HasFlag(Stage.EnrichRecords))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 3.5: Обогащение неполных записей (Geometry, Bolts, Welds) ===");
                    var enrichedCount = new JsonRecordEnricher().Enrich(merged);
                    Console.WriteLine($"Этап 3.5 завершён. Обогащено записей: {enrichedCount}");
                }

                var allJsonPath = Path.Combine(jsonAllDir, "all.json");
                JsonMerger.SaveToFile(merged, allJsonPath);
                Console.WriteLine($"  Записано {merged.Count} записей => {allJsonPath}");

                var allNotDuplicateJsonPath = Path.Combine(jsonAllDir, "all_NotDuplicate.json");

                // Этап 5: Создание списков profile.txt и CONNECTION_CODE.txt
                if (stages.HasFlag(Stage.ExportProfileAndConnectionCodeTxt))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 4: Создание списков profile.txt и CONNECTION_CODE.txt ===");
                    new ProfileAndConnectionCodeExporter().Export(allJsonPath, excelProfileOutDir);
                    Console.WriteLine("Этап 4 завершён.");
                }

                // Этап 6: Создание ProfileBeam.json и CONNECTION_CODE.json
                if (stages.HasFlag(Stage.ExportProfileAndConnectionCodeJson))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 5: Создание ProfileBeam.json и CONNECTION_CODE.json ===");
                    new TextListToJsonExporter().ExportProfileJson(
                        Path.Combine(excelProfileOutDir, "profile.txt"),
                        Path.Combine(excelProfileOutDir, "ProfileBeam.json"));
                    new TextListToJsonExporter().ExportConnectionCodeJson(
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE.txt"),
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE.json"));
                    Console.WriteLine("Этап 5 завершён.");
                }

                // Этап 7: Проверка all.json на дубликаты CONNECTION_CODE
                if (stages.HasFlag(Stage.CheckDuplicates))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 6: Проверка all.json на дубликаты CONNECTION_CODE ===");
                    var duplicates = new ConnectionCodeDuplicateChecker().FindDuplicates(
                        allJsonPath,
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE_duplicates.txt"));
                    Console.WriteLine($"Этап 6 завершён. Найдено дубликатов: {duplicates.Count}");
                }

                // Этап 8: Создание all_NotDuplicate.json с заменой дубликатов
                if (stages.HasFlag(Stage.CreateDeduplicatedJson))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 7: Создание all_NotDuplicate.json с заменой дубликатов ===");
                    var changedCodes = new ConnectionCodeDeduplicator().CreateDeduplicatedJson(
                        allJsonPath,
                        allNotDuplicateJsonPath,
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE_replacements.txt"));
                    Console.WriteLine($"Этап 7 завершён. Заменено CONNECTION_CODE: {changedCodes}");
                }

                // Этап 9: Создание CONNECTION_CODE_new.json и CONNECTION_CODE_new.txt
                if (stages.HasFlag(Stage.ExportNewConnectionCodes))
                {
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

                // Этап 10: Создание NameConnections.json
                if (stages.HasFlag(Stage.ExportNameConnections))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 9: Создание NameConnections.json из all_NotDuplicate.json ===");
                    new NameConnectionsExporter().Export(
                        allNotDuplicateJsonPath,
                        Path.Combine(excelProfileOutDir, "NameConnections.json"));
                    Console.WriteLine("Этап 9 завершён.");
                }

                // Этап 11: Копирование all_NotDuplicate.json в ConvertData.Data\JSON\
                if (stages.HasFlag(Stage.CopyToDataProject))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Этап 10: Копирование all_NotDuplicate.json в ConvertData.Data\\JSON\\ ===");
                    var dataJsonDir = Path.Combine(projectDir, "..", "ConvertData.Data", "JSON");
                    dataJsonDir = Path.GetFullPath(dataJsonDir);
                    Directory.CreateDirectory(dataJsonDir);
                    var destPath = Path.Combine(dataJsonDir, "all_NotDuplicate.json");
                    File.Copy(allNotDuplicateJsonPath, destPath, overwrite: true);
                    Console.WriteLine($"  Скопировано: {allNotDuplicateJsonPath} -> {destPath}");
                    Console.WriteLine("Этап 10 завершён.");
                }
            }

            // Этап 12: Экспорт анкеров из Anchor.xlsx в JSON
            if (stages.HasFlag(Stage.ExportAnchors))
            {
                Console.WriteLine();
                Console.WriteLine("=== Этап 11: Экспорт анкеров из Anchor.xlsx в JSON ===");
                Directory.CreateDirectory(excelAnchorOutDir);
                _anchorExcelExporter.Export(
                    excelAnchorDir,
                    Path.Combine(excelAnchorOutDir, "Anchor.json"));
                Console.WriteLine("Этап 11 завершён.");
            }

            // Этап 13: Экспорт анкеров из MarkSteel.xlsx в JSON
            if (stages.HasFlag(Stage.ExportMarkSteel))
            {
                Console.WriteLine();
                Console.WriteLine("=== Этап 12: Экспорт анкеров из MarkSteel.xlsx в JSON ===");
                Directory.CreateDirectory(exceSteelDirOut);
                _anchorExcelExporter.Export(
                    exceSteelDir,
                    Path.Combine(exceSteelDirOut, "MarkSteel.json"));
                Console.WriteLine("Этап 12 завершён.");
            }

            Console.WriteLine();
            Console.WriteLine("Все указанные этапы завершены.");
            Console.ReadKey();
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
                return args.Where(p => !string.Equals(Path.GetFileName(p), "ProfileBeam.xls", StringComparison.OrdinalIgnoreCase));
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
            Console.WriteLine("Создан json: " + outPath);
        }

        private void ApplyProfilesToJson(string jsonOutDir, string excelProfileDir)
        {
            var profileLookup = _profileLookupLoader.Load(excelProfileDir);
            if (profileLookup.Count == 0)
            {
                Console.WriteLine("ProfileBeam lookup is empty: EXCEL_Profile/ProfileBeam.xls was not parsed.");
                return;
            }
            _profilePatcher.SelfCheckProfile(profileLookup);
            _profilePatcher.ApplyProfilesToJson(jsonOutDir, profileLookup);
        }
    }
}
