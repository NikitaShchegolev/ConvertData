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
    internal sealed class XlsRowReader : IRowReader
    {
        public List<Row> Read(string path)
        {
            var rawLines = ReadLinesBestEffort(path);

            var lines = rawLines
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            if (lines.Count == 0)
                return new List<Row>();

            // Если это справочник профилей (TSV) с заголовком: Profile\tH\tB\ts\tt
            var headerParts = SplitByTab(lines[0]);
            bool isProfileTsv = headerParts.Count >= 5
                && string.Equals(NormalizeHeaderToken(headerParts[0]), "Profile", System.StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeHeaderToken(headerParts[1]), "H", System.StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeHeaderToken(headerParts[2]), "B", System.StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeHeaderToken(headerParts[3]), "s", System.StringComparison.OrdinalIgnoreCase)
                && string.Equals(NormalizeHeaderToken(headerParts[4]), "t", System.StringComparison.OrdinalIgnoreCase);

            if (isProfileTsv)
            {
                var list = new List<Row>();
                for (int i = 1; i < lines.Count; i++)
                {
                    var parts = SplitByTab(lines[i]);
                    var profile = Get(parts, 0, "");
                    if (string.IsNullOrWhiteSpace(profile))
                        continue;

                    var h = Get(parts, 1, "0");
                    var b = Get(parts, 2, "0");
                    var s = Get(parts, 3, "0");
                    var t = Get(parts, 4, "0");

                    list.Add(new Row
                    {
                        Profile = profile,
                        H = ParseDouble(h),
                        B = ParseDouble(b),
                        s = ParseDouble(s),
                        t = ParseDouble(t)
                    });
                }

                return list;
            }

            // Строим индекс колонок по заголовкам первой строки.
            var colIndex = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < headerParts.Count; c++)
            {
                var key = NormalizeHeaderToken(headerParts[c]);
                if (!string.IsNullOrEmpty(key) && !colIndex.ContainsKey(key))
                    colIndex[key] = c;
            }

            var rows = new List<Row>();

            for (int i = 1; i < lines.Count; i++)
            {
                var parts = SplitByTab(lines[i]);

                string name = GetByHeader(parts, colIndex, "Name", "");
                string code = GetByHeader(parts, colIndex, "CONNECTION_CODE", "");
                string profile = GetByHeader(parts, colIndex, "Profile", "");

                if (string.IsNullOrWhiteSpace(code) && string.IsNullOrWhiteSpace(name))
                    continue;

                string n = GetByHeader(parts, colIndex, "N", "0");
                string nt = GetByHeader(parts, colIndex, "Nt", n);
                string nc = GetByHeader(parts, colIndex, "Nc", n);
                string q = GetByHeader(parts, colIndex, "Q", "0");
                string qo = GetByHeader(parts, colIndex, "Qo", "0");
                string t = GetByHeader(parts, colIndex, "T", "0");
                string m = GetByHeader(parts, colIndex, "M", "0");
                string mneg = GetByHeader(parts, colIndex, "Mneg", "0");
                string mo = GetByHeader(parts, colIndex, "Mo", "0");
                string alpha = GetByHeader(parts, colIndex, "Alpha", "0");
                string beta = GetByHeader(parts, colIndex, "Beta", "0");
                string gamma = GetByHeader(parts, colIndex, "Gamma", "0");
                string delta = GetByHeader(parts, colIndex, "Delta", "0");
                string epsilon = GetByHeader(parts, colIndex, "Epsilon", "0");
                string lambda = GetByHeader(parts, colIndex, "Lambda", "0");

                rows.Add(Map15(name, code, profile, nt, nc, n, qo, q, t, m, mneg, mo, alpha, beta, gamma, delta, epsilon, lambda));
            }

            return rows;
        }

        private static IEnumerable<string> ReadLinesBestEffort(string path)
        {
            // Встречается, что "Profile.xls" — это CSV/TSV сохранённый в OEM866 или ANSI.
            // Пробуем несколько кодировок и выбираем ту, где распознаётся заголовок Profile/H/B/s/t.
            var encodings = new[]
            {
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                Encoding.UTF8,
                Encoding.GetEncoding(1251),
                Encoding.GetEncoding(866)
            };

            foreach (var enc in encodings)
            {
                try
                {
                    var lines = File.ReadAllLines(path, enc);
                    if (lines.Length == 0)
                        continue;

                    var headerParts = SplitByTab(lines[0]);
                    bool isProfileTsv = headerParts.Count >= 5
                        && string.Equals(NormalizeHeaderToken(headerParts[0]), "Profile", System.StringComparison.OrdinalIgnoreCase)
                        && string.Equals(NormalizeHeaderToken(headerParts[1]), "H", System.StringComparison.OrdinalIgnoreCase)
                        && string.Equals(NormalizeHeaderToken(headerParts[2]), "B", System.StringComparison.OrdinalIgnoreCase)
                        && string.Equals(NormalizeHeaderToken(headerParts[3]), "s", System.StringComparison.OrdinalIgnoreCase)
                        && string.Equals(NormalizeHeaderToken(headerParts[4]), "t", System.StringComparison.OrdinalIgnoreCase);

                    if (isProfileTsv)
                        return lines;

                    // Если это НЕ справочник профилей, но файл читается без ошибок — тоже возвращаем.
                    // (Дальше логика прочтёт как обычную TSV-таблицу из EXCEL.)
                    return lines;
                }
                catch
                {
                    // пробуем следующую
                }
            }

            // Последний шанс: как binary с UTF8 (как было раньше)
            return File.ReadAllLines(path, Encoding.UTF8);
        }

        private static string NormalizeHeaderToken(string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;

            var t = s.Trim();
            if (t.Length > 0 && t[0] == '\uFEFF')
                t = t.TrimStart('\uFEFF');
            return t;
        }

        private static string Get(List<string> parts, int index, string fallback)
        {
            if (index < 0 || index >= parts.Count)
                return fallback;

            var v = parts[index];
            return string.IsNullOrWhiteSpace(v) ? fallback : v;
        }

        private static string GetByHeader(List<string> parts, Dictionary<string, int> colIndex, string header, string fallback)
        {
            if (!colIndex.TryGetValue(header, out var index))
                return fallback;

            return Get(parts, index, fallback);
        }

        private static List<string> SplitByTab(string line)
        {
            return line.Split(new[] { '\t' }, System.StringSplitOptions.None)
                .Select(p => p.Trim())
                .ToList();
        }

        private static Row Map15(
            string name,
            string code,
            string profile,
            string nt,
            string nc,
            string n,
            string qo,
            string q,
            string t,
            string m,
            string mneg,
            string mo,
            string alpha,
            string beta,
            string gamma,
            string delta,
            string epsilon,
            string lambda)
        {
            int nInt = ParseInt(n);

            int ntInt = ParseInt(nt);
            if (ntInt == 0 && !string.IsNullOrWhiteSpace(n))
                ntInt = nInt;
            int ncInt = ParseInt(nc);
            if (ncInt == 0 && !string.IsNullOrWhiteSpace(n))
                ncInt = nInt;
            int qoInt = ParseInt(qo);
            int qInt = ParseInt(q);
            int tInt = ParseInt(t);
            int mInt = ParseInt(m);
            double mnegDouble = ParseDouble(mneg);
            double moDouble = ParseDouble(mo);
            double alphaDouble = ParseDouble(alpha);
            double betaDouble = ParseDouble(beta);
            double gammaDouble = ParseDouble(gamma);
            double deltaDouble = ParseDouble(delta);
            double epsilonDouble = ParseDouble(epsilon);
            double lambdaDouble = ParseDouble(lambda);

            return new Row
            {
                Name = name,
                CONNECTION_CODE = code,
                Profile = profile,
                Nt = ntInt,
                Nc = ncInt,
                N = nInt,
                Qo = qoInt,
                Q = qInt,
                T = tInt,
                M = mInt,
                Mneg = mnegDouble,
                Mo = moDouble,
                Alpha = alphaDouble,
                Beta = betaDouble,
                Gamma = gammaDouble,
                Delta = deltaDouble,
                Epsilon = epsilonDouble,
                Lambda = lambdaDouble
            };
        }

        private static double ParseDouble(string s)
        {
            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                return v;

            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, new CultureInfo("ru-RU"), out v))
                return v;

            return 0.0;
        }

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
