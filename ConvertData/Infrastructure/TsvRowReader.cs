using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ConvertData.Application;
using ConvertData.Domain;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Reader для TSV-подобных файлов.
    /// Используется для файлов, которые по расширению могут быть `.xls`, но фактически содержат табличный текст с табами.
    /// </summary>
    internal sealed class TsvRowReader : IRowReader
    {
        /// <summary>
        /// Читает файл как текст (UTF-8) и парсит строки, разделённые табуляцией.
        /// Первая строка считается заголовком и пропускается.
        /// </summary>
        /// <param name="path">Путь к входному файлу.</param>
        /// <returns>Список строк `Row`.</returns>
        public List<Row> Read(string path)
        {
            var lines = File.ReadAllLines(path, Encoding.UTF8)
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            if (lines.Count == 0)
                return new List<Row>();

            var rows = new List<Row>();
            for (int i = 1; i < lines.Count; i++)
            {
                var parts = SplitByTab(lines[i]);
                if (parts.Count < 7)
                    continue;

                rows.Add(MapBasic(parts[0], parts[1], parts[2], parts[3], parts[4], parts[5], parts[6]));
            }
            return rows;
        }

        /// <summary>
        /// Делит строку по табуляции с отбрасыванием пустых элементов.
        /// </summary>
        private static List<string> SplitByTab(string line)
        {
            return line.Split(new[] { '\t' }, System.StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => p.Length > 0)
                .ToList();
        }

        /// <summary>
        /// Преобразует набор строковых значений в доменную модель `Row`.
        /// </summary>
        private static Row MapBasic(string name, string code, string profile, string n, string q, string qo, string t)
        {
            var nInt = ParseInt(n);
            var qInt = ParseInt(q);
            var qoInt = ParseInt(qo);
            var tInt = ParseInt(t);

            return new Row
            {
                Name = name,
                CONNECTION_CODE = code,
                Profile = profile,
                Nt = nInt,
                Nc = nInt,
                N = nInt,
                Qo = qoInt,
                Q = qInt,
                T = tInt,
                M = 0,
                Mneg = 0.0,
                Mo = 0.0,
                Alpha = 0.0,
                Beta = 0.0,
                Gamma = 0.0,
                Delta = 0.0,
                Epsilon = 0.0,
                Lambda = 0.0
            };
        }

        /// <summary>
        /// Парсит целое число из строки (InvariantCulture / ru-RU), иначе возвращает 0.
        /// </summary>
        private static int ParseInt(string s)
        {
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                return v;

            if (int.TryParse(s, NumberStyles.Integer, new CultureInfo("ru-RU"), out v))
                return v;

            return 0;
        }
    }
}
