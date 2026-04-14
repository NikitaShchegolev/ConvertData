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
            Console.WriteLine("=== Интерактивный выбор блоков для выполнения ===");
            Console.WriteLine("Выберите номера блоков через запятую (например, 1,2,3) или введите 'all' для всех:");
            Console.WriteLine("  1. CreateJson - создание JSON из Excel");
            Console.WriteLine("  2. ApplyProfiles - применение справочника профилей");
            Console.WriteLine("  3. MergeAndEnrich - объединение и обогащение");
            Console.WriteLine("  4. ExportProfiles - экспорт профилей и кодов");
            Console.WriteLine("  5. Deduplication - дедупликация");
            Console.WriteLine("  6. CopyToData - копирование в Data проект");
            Console.WriteLine("  7. AnchorExport - экспорт анкеров из Anchor.xlsx");
            Console.WriteLine("  8. SteelExport - экспорт анкеров из MarkSteel.xlsx");
            Console.WriteLine("  9. Conversion - блок конвертации (1+2)");
            Console.WriteLine("  10. Processing - блок обработки (3+4+5+6)");
            Console.WriteLine("  11. Anchors - блок анкеров (7+8)");
            Console.WriteLine("  12. All - все блоки");
            Console.WriteLine();            
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

            // Если есть аргументы командной строки, выполняем один раз и выходим
            if (args.Length > 0 &&
                !args.Contains("--interactive", StringComparer.OrdinalIgnoreCase) &&
                !args.Contains("-i", StringComparer.OrdinalIgnoreCase))
            {
                // Режим командной строки: выполняем один раз
                var blocks = RunModeParser.GetBlocks(args);
                Console.WriteLine($"Выполняемые блоки: {blocks}");
                ExecuteBlocks(blocks, projectDir, excelDir, excelProfileDir, jsonOutDir, jsonAllDir,
                    excelProfileOutDir, excelAnchorDir, excelAnchorOutDir, exceSteelDir, exceSteelDirOut);
                Console.WriteLine();
                Console.WriteLine("Все указанные блоки завершены.");
                Console.WriteLine("Нажмите любую клавишу для выхода...");
                Console.ReadKey();
                return;
            }

            // Интерактивный режим с бесконечным циклом
            Console.WriteLine("=== Интерактивный режим ConvertData ===");
            Console.WriteLine("Введите 'exit' для выхода из программы.");
            Console.WriteLine();

            while (true)
            {
                var blocks = InteractiveBlockSelection();
                if (blocks == Block.None)
                {
                    Console.WriteLine("Не выбрано ни одного блока. Введите 'exit' для выхода или нажмите Enter для продолжения.");
                    var input = Console.ReadLine();
                    if (input?.Trim().Equals("exit", StringComparison.OrdinalIgnoreCase) == true)
                        break;
                    continue;
                }

                ExecuteBlocks(blocks, projectDir, excelDir, excelProfileDir, jsonOutDir, jsonAllDir,
                    excelProfileOutDir, excelAnchorDir, excelAnchorOutDir, exceSteelDir, exceSteelDirOut);

                Console.WriteLine();
                // После выполнения блоков сразу возвращаемся к выбору
            }

            Console.WriteLine("Программа завершена.");
        }

        private void ExecuteBlocks(Block blocks, string projectDir, string excelDir, string excelProfileDir,
            string jsonOutDir, string jsonAllDir, string excelProfileOutDir, string excelAnchorDir,
            string excelAnchorOutDir, string exceSteelDir, string exceSteelDirOut)
        {
            // Блок 1: CreateJson - создание JSON из Excel
            if (blocks.HasFlag(Block.CreateJson))
            {
                Console.WriteLine("=== Блок 1: Создание JSON из Excel (без профилей) ===");
                ClearJsonOut(jsonOutDir);

                foreach (var input in GetInputFiles(RunModeParser.GetInputArgsForCreateJson(Array.Empty<string>()), excelDir))
                    ConvertOne(input, jsonOutDir);

                Console.WriteLine("Блок 1 завершён.");
            }
            else if (blocks.HasFlag(Block.ApplyProfiles) || blocks.HasFlag(Block.MergeAndEnrich))
            {
                // Если пропущен блок CreateJson, но нужны последующие, проверяем наличие JSON файлов
                if (!Directory.Exists(jsonOutDir) || !Directory.EnumerateFiles(jsonOutDir, "*.json").Any())
                {
                    Console.WriteLine("ВНИМАНИЕ: Блок CreateJson пропущен, но в папке JSON_OUT нет JSON файлов.");
                    Console.WriteLine("Для выполнения последующих блоков нужны JSON файлы.");
                }
            }

            // Блок 2: ApplyProfiles - применение справочника профилей
            if (blocks.HasFlag(Block.ApplyProfiles))
            {
                Console.WriteLine();
                Console.WriteLine("=== Блок 2: Применение справочника профилей (Beam_H, Beam_B, Beam_s, Beam_t) ===");
                ApplyProfilesToJson(jsonOutDir, excelProfileDir);
                Console.WriteLine("Блок 2 завершён.");

                Console.WriteLine();
                Console.WriteLine("=== Дополнительно: Экспорт профилей из Excel → Profile.json ===");
                _profileExcelExporter.Export(
                    excelProfileDir,
                    Path.Combine(excelProfileOutDir, "Profile.json"));
                Console.WriteLine("Экспорт профилей завершён.");
            }

            // Блок 3: MergeAndEnrich - объединение и обогащение
            if (blocks.HasFlag(Block.MergeAndEnrich))
            {
                Console.WriteLine();
                Console.WriteLine("=== Блок 3: Объединение всех JSON в один файл ===");
                var merged = new JsonMerger().MergeAll(jsonOutDir);
                Console.WriteLine("Объединение завершено.");

                Console.WriteLine();
                Console.WriteLine("=== Блок 3.5: Обогащение неполных записей (Geometry, Bolts, Welds) ===");
                var enrichedCount = new JsonRecordEnricher().Enrich(merged);
                Console.WriteLine($"Обогащение завершено. Обогащено записей: {enrichedCount}");

                var allJsonPath = Path.Combine(jsonAllDir, "all.json");
                JsonMerger.SaveToFile(merged, allJsonPath);
                Console.WriteLine($"  Записано {merged.Count} записей => {allJsonPath}");

                // Сохраняем allJsonPath для использования в следующих блоках
                var allNotDuplicateJsonPath = Path.Combine(jsonAllDir, "all_NotDuplicate.json");

                // Блок 4: ExportProfiles - экспорт профилей и кодов
                if (blocks.HasFlag(Block.ExportProfiles))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Блок 4: Создание списков profile.txt и CONNECTION_CODE.txt ===");
                    new ProfileAndConnectionCodeExporter().Export(allJsonPath, excelProfileOutDir);
                    Console.WriteLine("Блок 4 завершён.");

                    Console.WriteLine();
                    Console.WriteLine("=== Блок 4.5: Создание ProfileBeam.json и CONNECTION_CODE.json ===");
                    new TextListToJsonExporter().ExportProfileJson(
                        Path.Combine(excelProfileOutDir, "profile.txt"),
                        Path.Combine(excelProfileOutDir, "ProfileBeam.json"));
                    new TextListToJsonExporter().ExportConnectionCodeJson(
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE.txt"),
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE.json"));
                    Console.WriteLine("Блок 4.5 завершён.");
                }

                // Блок 5: Deduplication - дедупликация
                if (blocks.HasFlag(Block.Deduplication))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Блок 5: Проверка all.json на дубликаты CONNECTION_CODE ===");
                    var duplicates = new ConnectionCodeDuplicateChecker().FindDuplicates(
                        allJsonPath,
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE_duplicates.txt"));
                    Console.WriteLine($"Проверка дубликатов завершена. Найдено дубликатов: {duplicates.Count}");

                    Console.WriteLine();
                    Console.WriteLine("=== Блок 5.5: Создание all_NotDuplicate.json с заменой дубликатов ===");
                    var changedCodes = new ConnectionCodeDeduplicator().CreateDeduplicatedJson(
                        allJsonPath,
                        allNotDuplicateJsonPath,
                        Path.Combine(excelProfileOutDir, "CONNECTION_CODE_replacements.txt"));
                    Console.WriteLine($"Дедупликация завершена. Заменено CONNECTION_CODE: {changedCodes}");

                    Console.WriteLine();
                    Console.WriteLine("=== Блок 5.7: Создание CONNECTION_CODE_new.json и CONNECTION_CODE_new.txt из all_NotDuplicate.json ===");
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
                    Console.WriteLine("Блок 5.7 завершён.");

                    Console.WriteLine();
                    Console.WriteLine("=== Блок 5.9: Создание NameConnections.json из all_NotDuplicate.json ===");
                    new NameConnectionsExporter().Export(
                        allNotDuplicateJsonPath,
                        Path.Combine(excelProfileOutDir, "NameConnections.json"));
                    Console.WriteLine("Блок 5.9 завершён.");
                }

                // Блок 6: CopyToData - копирование в Data проект
                if (blocks.HasFlag(Block.CopyToData))
                {
                    Console.WriteLine();
                    Console.WriteLine("=== Блок 6: Копирование all_NotDuplicate.json в ConvertData.Data\\JSON\\ ===");
                    var dataJsonDir = Path.Combine(projectDir, "..", "ConvertData.Data", "JSON");
                    dataJsonDir = Path.GetFullPath(dataJsonDir);
                    Directory.CreateDirectory(dataJsonDir);
                    var destPath = Path.Combine(dataJsonDir, "all_NotDuplicate.json");
                    File.Copy(allNotDuplicateJsonPath, destPath, overwrite: true);
                    Console.WriteLine($"  Скопировано: {allNotDuplicateJsonPath} -> {destPath}");
                    Console.WriteLine("Блок 6 завершён.");
                }
            }

            // Блок 7: AnchorExport - экспорт анкеров из Anchor.xlsx
            if (blocks.HasFlag(Block.AnchorExport))
            {
                Console.WriteLine();
                Console.WriteLine("=== Блок 7: Экспорт анкеров из Anchor.xlsx в JSON ===");
                Directory.CreateDirectory(excelAnchorOutDir);
                _anchorExcelExporter.Export(
                    excelAnchorDir,
                    Path.Combine(excelAnchorOutDir, "Anchor.json"));
                Console.WriteLine("Блок 7 завершён.");
            }

            // Блок 8: SteelExport - экспорт анкеров из MarkSteel.xlsx
            if (blocks.HasFlag(Block.SteelExport))
            {
                Console.WriteLine();
                Console.WriteLine("=== Блок 8: Экспорт анкеров из MarkSteel.xlsx в JSON ===");
                Directory.CreateDirectory(exceSteelDirOut);
                _anchorExcelExporter.Export(
                    exceSteelDir,
                    Path.Combine(exceSteelDirOut, "MarkSteel.json"));
                Console.WriteLine("Блок 8 завершён.");
            }
        }

        private Block InteractiveBlockSelection()
        {
            Console.Write("Ваш выбор: ");

            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input))
                return Block.None;

            if (input.Equals("all", StringComparison.OrdinalIgnoreCase))
                return Block.All;

            var blocks = Block.None;
            var parts = input.Split(',', ';', ' ')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s));

            foreach (var part in parts)
            {
                if (int.TryParse(part, out int number))
                {
                    var block = NumberToBlock(number);
                    if (block != Block.None)
                        blocks |= block;
                }
                else if (Enum.TryParse<Block>(part, true, out var namedBlock))
                {
                    blocks |= namedBlock;
                }
            }

            Console.WriteLine($"Выбраны блоки: {blocks}");
            return blocks;
        }

        private Block NumberToBlock(int number)
        {
            return number switch
            {
                1 => Block.CreateJson,
                2 => Block.ApplyProfiles,
                3 => Block.MergeAndEnrich,
                4 => Block.ExportProfiles,
                5 => Block.Deduplication,
                6 => Block.CopyToData,
                7 => Block.AnchorExport,
                8 => Block.SteelExport,
                9 => Block.Conversion,
                10 => Block.Processing,
                11 => Block.Anchors,
                12 => Block.All,
                _ => Block.None
            };
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
