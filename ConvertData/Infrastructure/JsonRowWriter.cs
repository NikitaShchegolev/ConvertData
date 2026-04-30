using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ConvertData.Application;
using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Записывает список объектов Row в JSON-файл с форматированием.
    /// </summary>
    internal sealed class JsonRowWriter : IRowWriter, IAsyncRowWriter
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

        private readonly ConnectionCodeDuplicateResolver _duplicateResolver = new();

        /// <summary>
        /// Записывает список объектов Row в JSON-файл.
        /// </summary>
        /// <param name="rows">Список объектов Row.</param>
        /// <param name="outputPath">Путь к выходному JSON-файлу.</param>
        public void Write(List<Row> rows, string outputPath)
        {
            // Обработка дубликатов CONNECTION_CODE
            var codes = rows.Select(r => r.CONNECTION_CODE).ToList();
            var processedCodes = _duplicateResolver.Process(codes);

            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.AppendLine("  {");
                sb.AppendLine("    \"Name\": \"" + JsonEscape(r.Name) + "\",");
                sb.AppendLine("    \"GUID_NODE\": \"" + Guid.NewGuid().ToString() + "\",");
                sb.AppendLine("    \"CONNECTION_CODE\": \"" + JsonEscape(processedCodes[i]) + "\",");
                sb.AppendLine("    \"TypeNode\": \"" + JsonEscape(r.TypeNode) + "\",");
                sb.AppendLine("    \"Gost\": \"" + JsonEscape(r.Gost) + "\",");
                sb.AppendLine("    \"variable\": \"" + JsonEscape(r.variable) + "\",");
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
                WriteBrace(sb, r);
                WriteRigel(sb, r);
                WriteRunThrough(sb, r);
                WritePlate(sb, r);
                WriteFlange(sb, r);
                WriteStiff(sb, r);
                WriteBase(sb, r);
                sb.AppendLine("    },");
                sb.AppendLine();

                // Bolts
                WriteBolts(sb, r);
                sb.AppendLine();

                // Welds
                WriteWelds(sb, r);
                sb.AppendLine();

                // Holes
                WriteHoles(sb, r);
                sb.AppendLine();

                // Anchor
                WriteAnchor(sb, r);
                sb.AppendLine();

                // Anchor
                WriteShearKey(sb, r);
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
            sb.AppendLine("        \"GostBeams\": \"" + JsonEscape(r.GostBeams) + "\",");
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
            sb.AppendLine("        \"Beam_i_z\": " + Dbl(r.Beam_iz) + ",");
            sb.AppendLine("        \"Beam_i_y\": " + Dbl(r.Beam_iy) + ",");
            sb.AppendLine("        \"Beam_xo\": " + Dbl(r.Beam_xo) + ",");
            sb.AppendLine("        \"Beam_yo\": " + Dbl(r.Beam_yo));
            sb.AppendLine("      },");
        }

        private static void WriteColumn(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Column\": {");
            sb.AppendLine("        \"ProfileColumn\": \"" + JsonEscape(r.ProfileColumn) + "\",");
            sb.AppendLine("        \"GostColumn\": \"" + JsonEscape(r.GostColumn) + "\",");
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
            sb.AppendLine("        \"Column_i_z\": " + Dbl(r.Column_iz) + ",");
            sb.AppendLine("        \"Column_i_y\": " + Dbl(r.Column_iy) + ",");
            sb.AppendLine("        \"Column_xo\": " + Dbl(r.Column_xo) + ",");
            sb.AppendLine("        \"Column_yo\": " + Dbl(r.Column_yo));
            sb.AppendLine("      },");
        }
        private void WriteBrace(StringBuilder sb, Row r)
        {

            sb.AppendLine("      \"Brace\": {");
            sb.AppendLine("        \"ProfileBrace\": \"" + JsonEscape(r.ProfileBrace) + "\",");
            sb.AppendLine("        \"GostBrace\": \"" + JsonEscape(r.GostBrace) + "\",");
            sb.AppendLine("        \"Brace_H\": " + Dbl(r.Brace_H) + ",");
            sb.AppendLine("        \"Brace_B\": " + Dbl(r.Brace_B) + ",");
            sb.AppendLine("        \"Brace_s\": " + Dbl(r.Brace_s) + ",");
            sb.AppendLine("        \"Brace_t\": " + Dbl(r.Brace_t) + ",");
            sb.AppendLine("        \"Brace_A\": " + Dbl(r.Brace_A) + ",");
            sb.AppendLine("        \"Brace_P\": " + Dbl(r.Brace_P) + ",");
            sb.AppendLine("        \"Brace_Iz\": " + Dbl(r.Brace_Iz) + ",");
            sb.AppendLine("        \"Brace_Iy\": " + Dbl(r.Brace_Iy) + ",");
            sb.AppendLine("        \"Brace_Ix\": " + Dbl(r.Brace_Ix) + ",");
            sb.AppendLine("        \"Brace_Wz\": " + Dbl(r.Brace_Wz) + ",");
            sb.AppendLine("        \"Brace_Wy\": " + Dbl(r.Brace_Wy) + ",");
            sb.AppendLine("        \"Brace_Wx\": " + Dbl(r.Brace_Wx) + ",");
            sb.AppendLine("        \"Brace_Sz\": " + Dbl(r.Brace_Sz) + ",");
            sb.AppendLine("        \"Brace_Sy\": " + Dbl(r.Brace_Sy) + ",");
            sb.AppendLine("        \"Brace_i_z\": " + Dbl(r.Brace_iz) + ",");
            sb.AppendLine("        \"Brace_i_y\": " + Dbl(r.Brace_iy) + ",");
            sb.AppendLine("        \"Brace_xo\": " + Dbl(r.Brace_xo) + ",");
            sb.AppendLine("        \"Brace_yo\": " + Dbl(r.Brace_yo) + ",");
            sb.AppendLine("        \"CountHoles\": {");
            sb.AppendLine("           \"e2\": " + Dbl(r.E2_Brace) + ",");
            sb.AppendLine("           \"e3\": " + Dbl(r.E3_Brace) + ",");
            sb.AppendLine("           \"n1\": " + Dbl(r.N1_Brace) + ",");
            sb.AppendLine("           \"n2\": " + Dbl(r.N2_Brace));
            sb.AppendLine("        }");
            sb.AppendLine("     },");
        }

        private static void WritePlate(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Plate\": {");
            sb.AppendLine("        \"H_Plate\": " + Dbl(r.H_Plate) + ",");
            sb.AppendLine("        \"B_Plate\": " + Dbl(r.B_Plate) + ",");
            sb.AppendLine("        \"Lws_Plate\": " + Dbl(r.Lws_Plate) + ",");
            sb.AppendLine("        \"tp_Plate\": " + Dbl(r.Tp_Plate) + ",");
            sb.AppendLine("        \"Plate_tr1\": " + Dbl(r.Tr1_Plate) + ",");
            sb.AppendLine("        \"Plate_tr2\": " + Dbl(r.Tr2_Plate));
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
            sb.AppendLine("        \"H\": " + Dbl(r.H_Stiff) + ",");
            sb.AppendLine("        \"B\": " + Dbl(r.B_Stiff) + ",");
            sb.AppendLine("        \"Lws\": " + Dbl(r.Lws_Stiff) + ",");
            sb.AppendLine("        \"tp\": " + Dbl(r.Tp_Stiff) + ",");
            sb.AppendLine("        \"tr1\": " + Dbl(r.Tr1_Stiff) + ",");
            sb.AppendLine("        \"tr2\": " + Dbl(r.Tr2_Stiff) + ",");
            sb.AppendLine("        \"Tg_Stiff\": " + Dbl(r.Tg_Stiff) + ",");
            sb.AppendLine("        \"Lg_Stiff\": " + Dbl(r.Lg_Stiff) + ",");
            sb.AppendLine("        \"Tf_Stiff\": " + Dbl(r.Tf_Stiff) + ",");
            sb.AppendLine("        \"Lh_Stiff\": " + Dbl(r.Lh_Stiff) + ",");
            sb.AppendLine("        \"Hh_Stiff\": " + Dbl(r.Hh_Stiff));
            sb.AppendLine("      },");
        }

        private static void WriteBase(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Base\": {");
            sb.AppendLine("        \"H_base\": " + Dbl(r.H_base) + ",");
            sb.AppendLine("        \"S_base\": " + Dbl(r.S_base) + ",");
            sb.AppendLine("        \"B_base\": " + Dbl(r.B_base) + ",");
            sb.AppendLine("        \"T_base\": " + Dbl(r.T_base) + ",");
            sb.AppendLine("        \"Lp_base\": " + Dbl(r.Lp_base) + ",");
            sb.AppendLine("        \"Ls_base\": " + Dbl(r.Ls_base) + ",");
            sb.AppendLine("        \"Lws_base\": " + Dbl(r.Lws_base) + ",");
            sb.AppendLine("        \"Tws_base\": " + Dbl(r.Tws_base));
            sb.AppendLine("      }");
        }

        private static void WriteBolts(StringBuilder sb, Row r)
        {
            sb.AppendLine("    \"Bolts\": {");
            sb.AppendLine("      \"GostBolts\": \"" + JsonEscape(r.GostBolts) + "\",");
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
        private static void WriteHoles(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Holes\": {");
            sb.AppendLine("        \"GostHoles\": \"" + JsonEscape(r.GostHoles) + "\",");
            sb.AppendLine("        \"OptionHoles\": " + Dbl(r.OptionHoles) + ",");
            sb.AppendLine("        \"TableBrandHoles\": \"" + JsonEscape(r.TableBrandHoles) + "\",");
            sb.AppendLine("        \"DiameterHolesForBolts\": " + r.F_holes + ",");
            sb.AppendLine("        \"Dws_holes\": " + Dbl(r.Dws_holes) + ",");
            sb.AppendLine("        \"Dp_holes\": " + Dbl(r.Dp_holes) + ",");
            sb.AppendLine("        \"CountHoles\": {");
            sb.AppendLine("          \"Nh_holes_1_4\": " + r.Nh_Holes_1_4 + ",");
            sb.AppendLine("          \"Nh_holes_1_8\": " + r.Nh_Holes_5_8 + ",");
            sb.AppendLine("          \"xh\": " + Dbl(r.Anchor_xh_holes));
            sb.AppendLine("        }");
            sb.AppendLine("      },");
        }
        private static void WriteAnchor(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Anchor\": {");
            sb.AppendLine("        \"GostAnchore\": \"" + JsonEscape(r.GostAnchore) + "\",");
            sb.AppendLine("        \"Anchor_Lws\": " + Dbl(r.Lws_base) + ",");
            sb.AppendLine("        \"Anchor_Lp_base\": " + Dbl(r.Lp_base) + ",");
            sb.AppendLine("        \"Anchor_Ls_base\": " + Dbl(r.Ls_base) + ",");
            sb.AppendLine("        \"Anchor_tws_base\": " + Dbl(r.Tws_base) + ",");
            sb.AppendLine("        \"Anchor_d_ws_base\": " + Dbl(r.D_ws_base) + ",");
            sb.AppendLine("        \"Anchor_d_p_base\": " + Dbl(r.D_p_base) + ",");
            sb.AppendLine("        \"Anchor_xh_base\": " + Dbl(r.Xh_base) + ",");
            sb.AppendLine("        \"Anchor_nh_base_var1\": " + Dbl(r.Nh_base_var1) + ",");
            sb.AppendLine("        \"Anchor_nh_base_var2\": " + Dbl(r.Nh_base_var2) + ",");
            sb.AppendLine("        \"Anchor_k_fws_base\": " + WeldValue(r.K_fws_base) + ",");
            sb.AppendLine("        \"Anchor_anchor_var_1\": \"" + JsonEscape(r.Anchor_var_1) + "\",");
            sb.AppendLine("        \"Anchor_anchor_var_2\": \"" + JsonEscape(r.Anchor_var_2) + "\",");
            sb.AppendLine("        \"Anchor_anchor_var_3\": \"" + JsonEscape(r.Anchor_var_3) + "\",");
            sb.AppendLine("        \"Anchor_anchor_var_4\": \"" + JsonEscape(r.Anchor_var_4) + "\"");
            sb.AppendLine("      },");
        }
        private void WriteShearKey(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"ShearKey\": {");
            sb.AppendLine("        \"Lp_shearKey\": " + Dbl(r.Lp_ShearKey) + ",");
            sb.AppendLine("        \"Ls_shearKey\": " + Dbl(r.Ls_ShearKey));
            sb.AppendLine("      },");
        }
        
        private static void WriteBoltX(StringBuilder sb, Row r)
        {
            sb.AppendLine("        \"X\": {");
            int d1 = r.d1 != 0 ? r.d1 : (r.CoordinatesBolts.Count > 0 ? r.CoordinatesBolts[0].X : 0);
            int d2 = r.d2 != 0 ? r.d2 : (r.CoordinatesBolts.Count > 1 ? r.CoordinatesBolts[1].X : 0);
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
            sb.AppendLine("      \"GostWeld\": \"" + JsonEscape(r.GostWeld) + "\",");
            sb.AppendLine("      \"kf1\": " + WeldValue(r.kf1) + ",");
            sb.AppendLine("      \"kf2\": " + WeldValue(r.kf2) + ",");
            sb.AppendLine("      \"kf3\": " + WeldValue(r.kf3) + ",");
            sb.AppendLine("      \"kf4\": " + WeldValue(r.kf4) + ",");
            sb.AppendLine("      \"kf5\": " + WeldValue(r.kf5) + ",");
            sb.AppendLine("      \"kf6\": " + WeldValue(r.kf6) + ",");
            sb.AppendLine("      \"kf7\": " + WeldValue(r.kf7) + ",");
            sb.AppendLine("      \"kf8\": " + WeldValue(r.kf8) + ",");
            sb.AppendLine("      \"kf9\": " + WeldValue(r.kf9) + ",");
            sb.AppendLine("      \"kf10\": " + WeldValue(r.kf10) + ",");
            sb.AppendLine("      \"Anchor_k_fws_base\": " + WeldValue(r.K_fws_base));
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
            sb.AppendLine("      \"N\": " + Dbl(r.N) + ",");
            sb.AppendLine("      \"Nt\": " + Dbl(r.Nt) + ",");
            sb.AppendLine("      \"Nc\": " + Dbl(r.Nc) + ",");
            sb.AppendLine("      \"My\": " + Dbl(r.My) + ",");
            sb.AppendLine("      \"My_compression\": " + Dbl(r.My_compression) + ",");
            sb.AppendLine("      \"My_tension\": " + Dbl(r.My_tension) + ",");
            sb.AppendLine("      \"Mz\": " + Dbl(r.Mz) + ",");
            sb.AppendLine("      \"Mz_tension\": " + Dbl(r.Mz_tension) + ",");
            sb.AppendLine("      \"Mz_compression\": " + Dbl(r.Mz_compression) + ",");
            sb.AppendLine("      \"Mx\": " + Dbl(r.Mx) + ",");
            sb.AppendLine("      \"Mw\": " + Dbl(r.Mw) + ",");
            sb.AppendLine("      \"Mneg\": " + Dbl(r.Mneg) + ",");
            sb.AppendLine("      \"T\": " + Dbl(r.T) + ",");
            sb.AppendLine("      \"Qy\": " + Dbl(r.Qy) + ",");
            sb.AppendLine("      \"Qz\": " + Dbl(r.Qz) + ",");
            sb.AppendLine("      \"Qx\": " + Dbl(r.Qx) + ",");
            sb.AppendLine("      \"F_base\": " + Dbl(r.F_base));
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

        /// <summary>
        /// Асинхронно записывает список объектов Row в JSON-файл с обработкой дубликатов CONNECTION_CODE.
        /// </summary>
        /// <param name="rows">Список объектов Row.</param>
        /// <param name="outputPath">Путь к выходному JSON-файлу.</param>
        public async Task WriteAsync(List<Row> rows, string outputPath)
        {
            // 1. Обработка дубликатов CONNECTION_CODE
            var codes = rows.Select(r => r.CONNECTION_CODE).ToList();
            var processedCodes = _duplicateResolver.Process(codes);

            // 2. Формирование JSON с паузой между строками
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.AppendLine("  {");
                sb.AppendLine("    \"Name\": \"" + JsonEscape(r.Name) + "\",");
                sb.AppendLine("    \"GUID_NODE\": \"" + Guid.NewGuid().ToString() + "\",");
                sb.AppendLine("    \"CONNECTION_CODE\": \"" + JsonEscape(processedCodes[i]) + "\",");
                sb.AppendLine("    \"TypeNode\": \"" + JsonEscape(r.TypeNode) + "\",");
                sb.AppendLine("    \"Gost\": \"" + JsonEscape(r.Gost) + "\",");
                sb.AppendLine("    \"variable\": \"" + JsonEscape(r.variable) + "\",");
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
                WriteBrace(sb, r); 
                WriteRigel(sb, r);
                WriteRunThrough(sb, r);
                WritePlate(sb, r);
                WriteFlange(sb, r);
                WriteStiff(sb, r);
                WriteBase(sb, r);
                sb.AppendLine("    },");
                sb.AppendLine();

                // Bolts
                WriteBolts(sb, r);
                sb.AppendLine();

                // Welds
                WriteWelds(sb, r);
                sb.AppendLine();

                // Holes
                WriteHoles(sb, r);
                sb.AppendLine();

                // Anchor
                WriteAnchor(sb, r);
                sb.AppendLine();

                // ShearKey
                WriteShearKey(sb, r);
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

                // Небольшая пауза между обработкой строк
                await Task.Delay(10);
            }
            sb.AppendLine("]");

            // 4. Асинхронная запись в файл
            await File.WriteAllTextAsync(outputPath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private void WriteRunThrough(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"RunThrough\": {");
            sb.AppendLine("        \"ProfileRunThrough\": \"" + JsonEscape(r.ProfileRunThrough) + "\",");
            sb.AppendLine("        \"GostRunThrough\": \"" + JsonEscape(r.   GostRunThrough) + "\",");
            sb.AppendLine("        \"RunThrough_H\": " + Dbl(r.              RunThrough_H) + ",");
            sb.AppendLine("        \"RunThrough_B\": " + Dbl(r.              RunThrough_B) + ",");
            sb.AppendLine("        \"RunThrough_s\": " + Dbl(r.              RunThrough_s) + ",");
            sb.AppendLine("        \"RunThrough_t\": " + Dbl(r.              RunThrough_t) + ",");
            sb.AppendLine("        \"RunThrough_A\": " + Dbl(r.              RunThrough_A) + ",");
            sb.AppendLine("        \"RunThrough_P\": " + Dbl(r.              RunThrough_P) + ",");
            sb.AppendLine("        \"RunThrough_Iz\": " + Dbl(r.             RunThrough_Iz) + ",");
            sb.AppendLine("        \"RunThrough_Iy\": " + Dbl(r.             RunThrough_Iy) + ",");
            sb.AppendLine("        \"RunThrough_Ix\": " + Dbl(r.             RunThrough_Ix) + ",");
            sb.AppendLine("        \"RunThrough_Wz\": " + Dbl(r.             RunThrough_Wz) + ",");
            sb.AppendLine("        \"RunThrough_Wy\": " + Dbl(r.             RunThrough_Wy) + ",");
            sb.AppendLine("        \"RunThrough_Wx\": " + Dbl(r.             RunThrough_Wx) + ",");
            sb.AppendLine("        \"RunThrough_Sz\": " + Dbl(r.             RunThrough_Sz) + ",");
            sb.AppendLine("        \"RunThrough_Sy\": " + Dbl(r.             RunThrough_Sy) + ",");
            sb.AppendLine("        \"RunThrough_i_z\": " + Dbl(r.            RunThrough_iz) + ",");
            sb.AppendLine("        \"RunThrough_i_y\": " + Dbl(r.            RunThrough_iy) + ",");
            sb.AppendLine("        \"RunThrough_xo\": " + Dbl(r.             RunThrough_xo) + ",");
            sb.AppendLine("        \"RunThrough_yo\": " + Dbl(r.             RunThrough_yo));
            sb.AppendLine("     },");
        }

        private void WriteRigel(StringBuilder sb, Row r)
        {
            sb.AppendLine("      \"Rigel\": {");
            sb.AppendLine("        \"ProfileRigel\": \"" + JsonEscape(r.ProfileRigel) + "\",");
            sb.AppendLine("        \"GostRigel\": \"" +    JsonEscape(r.GostRigel) + "\",");
            sb.AppendLine("        \"Rigel_H\": " +               Dbl(r.Rigel_H) + ",");
            sb.AppendLine("        \"Rigel_B\": " +               Dbl(r.Rigel_B) + ",");
            sb.AppendLine("        \"Rigel_s\": " +               Dbl(r.Rigel_s) + ",");
            sb.AppendLine("        \"Rigel_t\": " +               Dbl(r.Rigel_t) + ",");
            sb.AppendLine("        \"Rigel_A\": " +               Dbl(r.Rigel_A) + ",");
            sb.AppendLine("        \"Rigel_P\": " +               Dbl(r.Rigel_P) + ",");
            sb.AppendLine("        \"Rigel_Iz\": " +              Dbl(r.Rigel_Iz) + ",");
            sb.AppendLine("        \"Rigel_Iy\": " +              Dbl(r.Rigel_Iy) + ",");
            sb.AppendLine("        \"Rigel_Ix\": " +              Dbl(r.Rigel_Ix) + ",");
            sb.AppendLine("        \"Rigel_Wz\": " +              Dbl(r.Rigel_Wz) + ",");
            sb.AppendLine("        \"Rigel_Wy\": " +              Dbl(r.Rigel_Wy) + ",");
            sb.AppendLine("        \"Rigel_Wx\": " +              Dbl(r.Rigel_Wx) + ",");
            sb.AppendLine("        \"Rigel_Sz\": " +              Dbl(r.Rigel_Sz) + ",");
            sb.AppendLine("        \"Rigel_Sy\": " +              Dbl(r.Rigel_Sy) + ",");
            sb.AppendLine("        \"Rigel_i_z\": " +             Dbl(r.Rigel_iz) + ",");
            sb.AppendLine("        \"Rigel_i_y\": " +             Dbl(r.Rigel_iy) + ",");
            sb.AppendLine("        \"Rigel_xo\": " +              Dbl(r.Rigel_xo) + ",");
            sb.AppendLine("        \"Rigel_yo\": " +              Dbl(r.Rigel_yo));
            sb.AppendLine("     },");
        }
    }
}