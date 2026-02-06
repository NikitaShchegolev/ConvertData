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

            var rows = new List<Row>();

            // Первая строка — заголовок, дальше данные.
            for (int i = 1; i < lines.Count; i++)
            {
                var parts = SplitByTab(lines[i]);

                string name = Get(parts, 0, "");
                string code = Get(parts, 1, "");
                string profile = Get(parts, 2, "");

                // Условие: в выход должны попадать только заполненные строки (по CONNECTION_CODE).
                if (string.IsNullOrWhiteSpace(code))
                    continue;

                string nt = Get(parts, 3, "0");
                string nc = Get(parts, 4, "0");
                string n = Get(parts, 5, "0");
                string qo = Get(parts, 6, "0");
                string q = Get(parts, 7, "0");
                string t = Get(parts, 8, "0");
                string m = Get(parts, 9, "0");
                string mneg = Get(parts, 10, "0");
                string mo = Get(parts, 11, "0");
                string alpha = Get(parts, 12, "0");
                string beta = Get(parts, 13, "0");
                string gamma = Get(parts, 14, "0");
                string delta = Get(parts, 15, "0");
                string epsilon = Get(parts, 16, "0");
                string lambda = Get(parts, 17, "0");

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
