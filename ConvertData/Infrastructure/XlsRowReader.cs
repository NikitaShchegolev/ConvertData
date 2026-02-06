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
        private const int DataRowCount = 15;

        public List<Row> Read(string path)
        {
            var lines = File.ReadAllLines(path, Encoding.UTF8)
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            if (lines.Count == 0)
                return new List<Row>();

            var rows = new List<Row>(capacity: DataRowCount);

            int dataLinesCount = lines.Count - 1;
            int take = dataLinesCount > 0 ? System.Math.Min(DataRowCount, dataLinesCount) : 0;

            for (int i = 0; i < take; i++)
            {
                var parts = SplitByTab(lines[i + 1]);

                string name = Get(parts, 0, "");
                string code = Get(parts, 1, "");
                string profile = Get(parts, 2, "");

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

            while (rows.Count < DataRowCount)
                rows.Add(new Row());

            return rows;
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
