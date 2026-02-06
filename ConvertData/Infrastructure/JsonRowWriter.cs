using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ConvertData.Application;
using ConvertData.Domain;

namespace ConvertData.Infrastructure
{
    internal sealed class JsonRowWriter : IRowWriter
    {
        public void Write(List<Row> rows, string outputPath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.AppendLine("  {");
                sb.AppendLine("    \"Name\": \"" + JsonEscape(r.Name) + "\",");
                sb.AppendLine("    \"CONNECTION_CODE\": \"" + JsonEscape(r.CONNECTION_CODE) + "\",");
                sb.AppendLine("    \"Profile\": \"" + JsonEscape(r.Profile) + "\",");
                sb.AppendLine("    \"Nt\": " + r.Nt + ",");
                sb.AppendLine("    \"Q\": " + r.Q + ",");
                sb.AppendLine("    \"Qo\": " + r.Qo + ",");
                sb.AppendLine("    \"T\": " + r.T + ",");
                sb.AppendLine("    \"Nc\": " + r.Nc + ",");
                sb.AppendLine("    \"N\": " + r.N + ",");
                sb.AppendLine("    \"M\": " + r.M + ",");
                sb.AppendLine("    \"Mneg\": " + r.Mneg.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"Mo\": " + r.Mo.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"α\": " + r.Alpha.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"β\": " + r.Beta.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"γ\": " + r.Gamma.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"δ\": " + r.Delta.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"ε\": " + r.Epsilon.ToString(CultureInfo.InvariantCulture) + ",");
                sb.AppendLine("    \"λ\": " + r.Lambda.ToString(CultureInfo.InvariantCulture));
                sb.Append("  }");
                if (i != rows.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");

            // Явно сохраняем UTF-8 без BOM.
            File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        /// <summary>
        /// Экранирует строку для безопасной вставки в JSON.
        /// </summary>
        /// <param name="s">Исходная строка.</param>
        /// <returns>Экранированная строка.</returns>
        private static string JsonEscape(string s)
        {
            if (s == null) return "";
            var sb = new StringBuilder(s.Length + 16);
            foreach (var ch in s)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 32)
                            sb.Append("\\u" + ((int)ch).ToString("x4"));
                        else
                            sb.Append(ch);
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
