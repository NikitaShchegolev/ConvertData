using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ConvertData.Application;
using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Записывает список объектов Row в JSON-файл с форматированием.
    /// </summary>
    internal sealed class JsonRowWriter : IRowWriter
    {
        /// <summary>
        /// Ключи для координат Y болтов (e1, p1-p10).
        /// </summary>
        private static readonly string[] BoltYKeys =
        [
            "Bolt1_e1", "Bolt2_p1", "Bolt3_p2", "Bolt4_p3", "Bolt5_p4",
            "Bolt6_p5", "Bolt7_p6", "Bolt8_p7", "Bolt9_p8", "Bolt10_p9", "Bolt11_p10"
        ];

        /// <summary>
        /// Ключи для координат X болтов (d1, d2).
        /// </summary>
        private static readonly string[] BoltXKeys = ["d1", "d2"];

        /// <summary>
        /// Записывает список объектов Row в JSON-файл.
        /// </summary>
        /// <param name="rows">Список объектов Row.</param>
        /// <param name="outputPath">Путь к выходному JSON-файлу.</param>
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
                sb.AppendLine("    \"TypeNode\": \"" + JsonEscape(r.TypeNode) + "\",");
                sb.AppendLine("    \"variable\": " + r.variable + ",");
                sb.AppendLine("    \"TableBrand\": \"" + JsonEscape(r.TableBrand) + "\",");
                sb.AppendLine("    \"Explanations\": \"" + JsonEscape(r.Explanations) + "\",");
                sb.AppendLine();

                // Stiffness
                sb.AppendLine("    \"Stiffness\": {");
                sb.AppendLine("      \"Sj\": " + r.Sj + ",");
                sb.AppendLine("      \"Sjo\": " + r.Sjo);
                sb.AppendLine("    },");
                sb.AppendLine();

                // Geometry
                sb.AppendLine("    \"Geometry\": {");
                WriteBeam(sb, r);
                WriteColumn(sb, r);
                WritePlate(sb, r);
                WriteFlange(sb, r);
                WriteStiff(sb, r);
                sb.AppendLine("    },");
                sb.AppendLine();

                // Bolts
                WriteBolts(sb, r);
                sb.AppendLine();

                // Welds
                WriteWelds(sb, r);
                sb.AppendLine();

                // InternalForces
                WriteInternalForces(sb, r);
                sb.AppendLine();

                // Coefficients
                sb.AppendLine("    \"Coefficients\": {");
                sb.AppendLine("      \"Alpha\": " + Dbl(r.Alpha) + ",");
                sb.AppendLine("      \"Beta\": " + Dbl(r.Beta) + ",");
                sb.AppendLine("      \"Gamma\": " + Dbl(r.Gamma) + ",");
                sb.AppendLine("      \"Delta\": " + Dbl(r.Delta) + ",");
                sb.AppendLine("      \"Epsilon\": " + Dbl(r.Epsilon) + ",");
                sb.AppendLine("      \"Lambda\": " + Dbl(r.Lambda));
                sb.AppendLine("    }");

                sb.Append("  }");
                if (i != rows.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");

            File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private static void WriteBeam(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Beam\": {");
            sb.AppendLine("        \"ProfileBeam\": \"" + JsonEscape(r.ProfileBeam) + "\",");
            sb.AppendLine("        \"Beam_H\": " + Dbl(r.Beam_H) + ",");
            sb.AppendLine("        \"Beam_B\": " + Dbl(r.Beam_B) + ",");
            sb.AppendLine("        \"Beam_s\": " + Dbl(r.Beam_s) + ",");
            sb.AppendLine("        \"Beam_t\": " + Dbl(r.Beam_t) + ",");
            sb.AppendLine("        \"Beam_A\": " + Dbl(r.Beam_A) + ",");
            sb.AppendLine("        \"Beam_P\": " + Dbl(r.Beam_P) + ",");
            sb.AppendLine("        \"Beam_Iz\": " + Dbl(r.Beam_Iz) + ",");
            sb.AppendLine("        \"Beam_Iy\": " + Dbl(r.Beam_Iy) + ",");
            sb.AppendLine("        \"Beam_Ix\": " + Dbl(r.Beam_Ix) + ",");
            sb.AppendLine("        \"Beam_Wz\": " + Dbl(r.Beam_Wz) + ",");
            sb.AppendLine("        \"Beam_Wy\": " + Dbl(r.Beam_Wy) + ",");
            sb.AppendLine("        \"Beam_Wx\": " + Dbl(r.Beam_Wx) + ",");
            sb.AppendLine("        \"Beam_Sz\": " + Dbl(r.Beam_Sz) + ",");
            sb.AppendLine("        \"Beam_Sy\": " + Dbl(r.Beam_Sy) + ",");
            sb.AppendLine("        \"Beam_iz\": " + Dbl(r.Beam_iz) + ",");
            sb.AppendLine("        \"Beam_iy\": " + Dbl(r.Beam_iy) + ",");
            sb.AppendLine("        \"Beam_xo\": " + Dbl(r.Beam_xo) + ",");
            sb.AppendLine("        \"Beam_yo\": " + Dbl(r.Beam_yo));
            sb.AppendLine("      },");
        }

        private static void WriteColumn(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Column\": {");
            sb.AppendLine("        \"ProfileColumn\": \"" + JsonEscape(r.ProfileColumn) + "\",");
            sb.AppendLine("        \"Column_H\": " + Dbl(r.Column_H) + ",");
            sb.AppendLine("        \"Column_B\": " + Dbl(r.Column_B) + ",");
            sb.AppendLine("        \"Column_s\": " + Dbl(r.Column_s) + ",");
            sb.AppendLine("        \"Column_t\": " + Dbl(r.Column_t) + ",");
            sb.AppendLine("        \"Column_A\": " + Dbl(r.Column_A) + ",");
            sb.AppendLine("        \"Column_P\": " + Dbl(r.Column_P) + ",");
            sb.AppendLine("        \"Column_Iz\": " + Dbl(r.Column_Iz) + ",");
            sb.AppendLine("        \"Column_Iy\": " + Dbl(r.Column_Iy) + ",");
            sb.AppendLine("        \"Column_Ix\": " + Dbl(r.Column_Ix) + ",");
            sb.AppendLine("        \"Column_Wz\": " + Dbl(r.Column_Wz) + ",");
            sb.AppendLine("        \"Column_Wy\": " + Dbl(r.Column_Wy) + ",");
            sb.AppendLine("        \"Column_Wx\": " + Dbl(r.Column_Wx) + ",");
            sb.AppendLine("        \"Column_Sz\": " + Dbl(r.Column_Sz) + ",");
            sb.AppendLine("        \"Column_Sy\": " + Dbl(r.Column_Sy) + ",");
            sb.AppendLine("        \"Column_iz\": " + Dbl(r.Column_iz) + ",");
            sb.AppendLine("        \"Column_iy\": " + Dbl(r.Column_iy) + ",");
            sb.AppendLine("        \"Column_xo\": " + Dbl(r.Column_xo) + ",");
            sb.AppendLine("        \"Column_yo\": " + Dbl(r.Column_yo));
            sb.AppendLine("      },");
        }

        private static void WritePlate(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Plate\": {");
            sb.AppendLine("        \"Plate_H\": " + Dbl(r.Plate_H) + ",");
            sb.AppendLine("        \"Plate_B\": " + Dbl(r.Plate_B) + ",");
            sb.AppendLine("        \"Plate_t\": " + Dbl(r.Plate_t));
            sb.AppendLine("      },");
        }

        private static void WriteFlange(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Flange\": {");
            sb.AppendLine("        \"Flange_Lb\": " + Dbl(r.Flange_Lb) + ",");
            sb.AppendLine("        \"Flange_H\": " + Dbl(r.Flange_H) + ",");
            sb.AppendLine("        \"Flange_B\": " + Dbl(r.Flange_B) + ",");
            sb.AppendLine("        \"Flange_t\": " + Dbl(r.Flange_t));
            sb.AppendLine("      },");
        }

        private static void WriteStiff(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Stiff\": {");
            sb.AppendLine("        \"Stiff_tbp\": " + Dbl(r.Stiff_tbp) + ",");
            sb.AppendLine("        \"Stiff_tg\": " + Dbl(r.Stiff_tg) + ",");
            sb.AppendLine("        \"Stiff_tf\": " + Dbl(r.Stiff_tf) + ",");
            sb.AppendLine("        \"Stiff_Lh\": " + Dbl(r.Stiff_Lh) + ",");
            sb.AppendLine("        \"Stiff_Hh\": " + Dbl(r.Stiff_Hh) + ",");
            sb.AppendLine("        \"Stiff_tr1\": " + Dbl(r.Stiff_tr1) + ",");
            sb.AppendLine("        \"Stiff_tr2\": " + Dbl(r.Stiff_tr2) + ",");
            sb.AppendLine("        \"Stiff_twp\": " + Dbl(r.Stiff_twp));
            sb.AppendLine("      }");
        }

        private static void WriteBolts(StringBuilder sb, Row r)
        {
            sb.AppendLine("    \"Bolts\": {");
            sb.AppendLine("      \"Option\": {");
            sb.AppendLine("        \"version\": " + r.OptionBolts);
            sb.AppendLine("      },");
            sb.AppendLine("      \"DiameterBolt\": {");
            sb.AppendLine("        \"F\": " + r.F);
            sb.AppendLine("      },");
            sb.AppendLine("      \"CountBolt\": {");
            sb.AppendLine("        \"Bolts_Nb\": " + r.Bolts_Nb);
            sb.AppendLine("      },");
            sb.AppendLine("      \"BoltRow\": {");
            sb.AppendLine("        \"N_Rows\": " + r.N_Rows);
            sb.AppendLine("      },");
            sb.AppendLine("      \"CoordinatesBolts\": {");
            WriteBoltY(sb, r);
            WriteBoltX(sb, r);
            WriteBoltZ(sb, r);
            sb.AppendLine("      }");
            sb.AppendLine("    },");
        }

        private static void WriteBoltY(StringBuilder sb, Row r)
        {
            sb.AppendLine("        \"Y\": {");
            sb.AppendLine("          \"Bolt1_e1\": " + r.e1 + ",");
            sb.AppendLine("          \"Bolt2_p1\": " + Dbl(r.p1) + ",");
            sb.AppendLine("          \"Bolt3_p2\": " + Dbl(r.p2) + ",");
            sb.AppendLine("          \"Bolt4_p3\": " + Dbl(r.p3) + ",");
            sb.AppendLine("          \"Bolt5_p4\": " + Dbl(r.p4) + ",");
            sb.AppendLine("          \"Bolt6_p5\": " + Dbl(r.p5) + ",");
            sb.AppendLine("          \"Bolt7_p6\": " + Dbl(r.p6) + ",");
            sb.AppendLine("          \"Bolt8_p7\": " + Dbl(r.p7) + ",");
            sb.AppendLine("          \"Bolt9_p8\": " + Dbl(r.p8) + ",");
            sb.AppendLine("          \"Bolt10_p9\": " + Dbl(r.p9) + ",");
            sb.AppendLine("          \"Bolt11_p10\": " + Dbl(r.p10));
            sb.AppendLine("        },");
        }

        private static void WriteBoltX(StringBuilder sb, Row r)
        {
            sb.AppendLine("        \"X\": {");
            int d1 = r.CoordinatesBolts.Count > 0 ? r.CoordinatesBolts[0].X : 0;
            int d2 = r.CoordinatesBolts.Count > 1 ? r.CoordinatesBolts[1].X : 0;
            sb.AppendLine("          \"d1\": " + d1 + ",");
            sb.AppendLine("          \"d2\": " + d2);
            sb.AppendLine("        },");
        }

        private static void WriteBoltZ(StringBuilder sb, Row r)
        {
            int val = r.CoordinatesBolts.Count > 0 ? r.CoordinatesBolts[0].Z : 0;
            sb.AppendLine("        \"Z\": {");
            sb.AppendLine("          \"BoltCoordinateZ\": " + val);
            sb.AppendLine("        }");
        }

        private static void WriteWelds(StringBuilder sb, Row r)
        {
            sb.AppendLine("    \"Welds\": {");
            sb.AppendLine("      \"kf1\": " + WeldValue(r.kf1) + ",");
            sb.AppendLine("      \"kf2\": " + WeldValue(r.kf2) + ",");
            sb.AppendLine("      \"kf3\": " + WeldValue(r.kf3) + ",");
            sb.AppendLine("      \"kf4\": " + WeldValue(r.kf4) + ",");
            sb.AppendLine("      \"kf5\": " + WeldValue(r.kf5) + ",");
            sb.AppendLine("      \"kf6\": " + WeldValue(r.kf6) + ",");
            sb.AppendLine("      \"kf7\": " + WeldValue(r.kf7) + ",");
            sb.AppendLine("      \"kf8\": " + WeldValue(r.kf8) + ",");
            sb.AppendLine("      \"kf9\": " + WeldValue(r.kf9) + ",");
            sb.AppendLine("      \"kf10\": " + WeldValue(r.kf10));
            sb.AppendLine("    },");
        }
        /// <summary>
        /// Метод для обработки значения сварки. 
        /// Если значение может быть преобразовано в число, 
        /// оно записывается как число. В противном случае
        /// оно записывается как строка с экранированием.
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        private static string WeldValue(string v)
        {
            if (string.IsNullOrWhiteSpace(v))
                return "\"\"";

            var value = v.Trim();
            var numeric = NumericParser.ParseDouble(value);

            if (IsNumericValue(value, numeric))
                return numeric.ToString(CultureInfo.InvariantCulture);

            return "\"" + JsonEscape(value) + "\"";
        }

        private static bool IsNumericValue(string source, double parsed)
        {
            if (double.TryParse(source, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _))
                return true;

            if (double.TryParse(source, NumberStyles.Float | NumberStyles.AllowThousands, new CultureInfo("ru-RU"), out _))
                return true;

            return false;
        }

        private static void WriteInternalForces(StringBuilder sb, Row r)
        {
            sb.AppendLine("    \"InternalForces\": {");
            sb.AppendLine("      \"N\": " + r.N + ",");
            sb.AppendLine("      \"Nt\": " + r.Nt + ",");
            sb.AppendLine("      \"Nc\": " + r.Nc + ",");
            sb.AppendLine("      \"My\": " + r.My + ",");
            sb.AppendLine("      \"Mz\": " + Dbl(r.Mz) + ",");
            sb.AppendLine("      \"Mx\": " + Dbl(r.Mx) + ",");
            sb.AppendLine("      \"Mw\": " + Dbl(r.Mw) + ",");
            sb.AppendLine("      \"Mneg\": " + Dbl(r.Mneg) + ",");
            sb.AppendLine("      \"T\": " + r.T + ",");
            sb.AppendLine("      \"Qy\": " + r.Qy + ",");
            sb.AppendLine("      \"Qz\": " + r.Qz + ",");
            sb.AppendLine("      \"Qx\": " + r.Qx);
            sb.AppendLine("    },");
        }
        /// <summary>
        /// Метод для преобразования double в строку
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        private static string Dbl(double v) => v.ToString(CultureInfo.InvariantCulture);

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
                    case '\t': sb.Append("\\t_f"); break;
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
