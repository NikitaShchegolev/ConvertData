using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ConvertData.Domain;
using ConvertData.Infrastructure;

namespace ConvertData.Application
{
    /// <summary>
    /// Сценарий (use-case) приложения: конвертация входных Excel/табличных файлов в JSON.
    ///
    /// Ответственность класса:
    /// - настроить окружение (кодировки, лицензия EPPlus);
    /// - определить папки ввода/вывода относительно каталога проекта;
    /// - найти входные файлы (или взять из аргументов);
    /// - для каждого входного файла выбрать подходящий reader и сохранить JSON.
    /// </summary>
    internal sealed class ConvertApp
    {
        private readonly IRowWriter _writer;
        private readonly IRowReaderFactory _readerFactory;
        private readonly IPathResolver _pathResolver;
        private readonly ILicenseConfigurator _licenseConfigurator;

        /// <summary>
        /// Создаёт экземпляр приложения с конкретными инфраструктурными реализациями.
        /// </summary>
        public ConvertApp()
        {
            _writer = new JsonRowWriter();
            _readerFactory = new RowReaderFactory();
            _pathResolver = new PathResolver();
            _licenseConfigurator = new EpplusLicenseConfigurator();
        }

        /// <summary>
        /// Основной запуск сценария конвертации.
        /// </summary>
        /// <param name="args">
        /// Аргументы командной строки.
        /// Если указаны — каждый аргумент считается путём к входному файлу и обрабатывается.
        /// Если не указаны — обрабатываются все поддерживаемые файлы из `EXCEL`.
        /// </param>
        public void Run(string[] args)
        {
            // Нужен для поддержки legacy-кодировок (например Windows-1251) на .NET.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Настройка лицензии библиотеки EPPlus.
            _licenseConfigurator.Configure();

            // Определяем папку проекта и папки ввода/вывода.
            var projectDir = _pathResolver.GetProjectDir(AppDomain.CurrentDomain.BaseDirectory) ?? AppDomain.CurrentDomain.BaseDirectory;
            var excelDir = Path.Combine(projectDir, "EXCEL");
            var jsonOutDir = Path.Combine(projectDir, "JSON_OUT");
            Directory.CreateDirectory(jsonOutDir);

            foreach (var f in Directory.EnumerateFiles(jsonOutDir, "*.json", SearchOption.TopDirectoryOnly))
            {
                try { File.Delete(f); } catch { }
            }

            // Если переданы аргументы — обрабатываем только их.
            if (args.Length > 0)
            {
                foreach (var input in args)
                    ConvertOne(input, jsonOutDir);

                return;
            }

            // Иначе обрабатываем всё из папки EXCEL.
            if (!Directory.Exists(excelDir))
                throw new DirectoryNotFoundException("Input folder not found: " + excelDir);

            var inputFiles = Directory.EnumerateFiles(excelDir)
                .Where(PathResolver.HasExcelExtension)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var input in inputFiles)
                ConvertOne(input, jsonOutDir);
        }

        /// <summary>
        /// Конвертирует один входной файл в JSON.
        /// </summary>
        /// <param name="inputPath">Путь к входному файлу.</param>
        /// <param name="jsonOutDir">Папка, куда сохраняется итоговый JSON.</param>
        private void ConvertOne(string inputPath, string jsonOutDir)
        {
            if (!File.Exists(inputPath))
                return;

            // Подбираем нужный reader (например TSV или Excel).
            var reader = _readerFactory.Create(inputPath);
            var rows = reader.Read(inputPath);

            // Пишем JSON с таким же базовым именем файла.
            var outPath = Path.Combine(jsonOutDir, Path.GetFileNameWithoutExtension(inputPath) + ".json");
            _writer.Write(rows, outPath);
            Console.WriteLine("Written: " + outPath);
        }
    }
}
