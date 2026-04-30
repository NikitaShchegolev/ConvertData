using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

/// <summary>
/// Отображает строковые данные из Excel в объекты Row.
/// Преобразует текстовые значения в соответствующие числовые типы.
/// </summary>
internal static class RowMapper
{
    internal static Row MapMainRowIdentity(
        string name,
        string code,
        string typeNode,
        string gost,
        string gostColumnAndBeams,
        string gostHoles,
        string gostBolts,
        string gostAnchore,
        string gostWeld,
        string gostProfile,
        string tableBrand,
        string profileBeam,
        string profileColumn,
        string explanations)
    {
        return new Row
        {
            Name = name,
            CONNECTION_CODE = code,
            TypeNode = typeNode ?? "",
            Gost = gost ?? "",
            GostBeams = gostColumnAndBeams ?? "",
            GostHoles = gostHoles ?? "",
            GostBolts = gostBolts ?? "",
            GostAnchore = gostAnchore ?? "",
            GostWeld = gostWeld ?? "",
            GostColumn = gostProfile ?? "",
            TableBrand = tableBrand ?? "",
            ProfileBeam = profileBeam ?? "",
            ProfileColumn = profileColumn ?? "",
            Explanations = explanations ?? "",
        };
    }

    internal static void MapMainRowStiffness(
        Row row,
        string variable,
        string sj,
        string sjo)
    {
        row.variable = variable ?? "";
        row.Sj = NumericParser.ParseInt(sj);
        row.Sjo = NumericParser.ParseInt(sjo);
    }

    internal static void MapMainRowForces(
        Row row,
        string nt,
        string q,
        string qz,
        string t,
        string nc,
        string n,
        string my,
        string my_compression,
        string my_tension,
        string mneg,
        string mz,
        string mz_compression,
        string mz_tension,
        string mx,
        string mw)
    {
        row.Nt = NumericParser.ParseInt(nt);
        row.Nc = NumericParser.ParseInt(nc);
        row.N = NumericParser.ParseInt(n);
        row.Qz = NumericParser.ParseInt(qz);
        row.Qy = NumericParser.ParseInt(q);
        row.My = NumericParser.ParseInt(my);
        row.My_compression = NumericParser.ParseInt(my_compression);
        row.My_tension = NumericParser.ParseInt(my_tension);
        row.Mz = NumericParser.ParseDouble(mz);
        row.Mz_compression = NumericParser.ParseDouble(mz_compression);
        row.Mz_tension = NumericParser.ParseDouble(mz_tension);
        row.Mx = NumericParser.ParseDouble(mx);
        row.Mw = NumericParser.ParseDouble(mw);
        row.T = NumericParser.ParseInt(t);
        row.Mneg = NumericParser.ParseDouble(mneg);
    }

    internal static void MapMainRowCoefficients(
        Row row,
        string alpha,
        string beta,
        string gamma,
        string delta,
        string epsilon,
        string lambda)
    {
        row.Alpha = NumericParser.ParseDouble(alpha);
        row.Beta = NumericParser.ParseDouble(beta);
        row.Gamma = NumericParser.ParseDouble(gamma);
        row.Delta = NumericParser.ParseDouble(delta);
        row.Epsilon = NumericParser.ParseDouble(epsilon);
        row.Lambda = NumericParser.ParseDouble(lambda);
    }

    internal static void MapMainRowBeamGeometry(
        Row row,
        string h,
        string b,
        string s,
        string tGeom)
    {
        row.Beam_H = NumericParser.ParseDouble(h);
        row.Beam_B = NumericParser.ParseDouble(b);
        row.Beam_s = NumericParser.ParseDouble(s);
        row.Beam_t = NumericParser.ParseDouble(tGeom);
    }

    internal static void MapMainRowPlateGeometry(
        Row row,
        string plateWidth,
        string plateHeight,
        string plateWeldLength,
        string plateThickness,
        string plateChamfer1,
        string plateChamfer2)
    {
        row.B_Plate = NumericParser.ParseDouble(plateWidth);
        row.H_Plate = NumericParser.ParseDouble(plateHeight);
        row.Lws_Plate = NumericParser.ParseDouble(plateWeldLength);
        row.Tp_Plate = NumericParser.ParseDouble(plateThickness);
        row.Tr1_Plate = NumericParser.ParseDouble(plateChamfer1);
        row.Tr2_Plate = NumericParser.ParseDouble(plateChamfer2);
    }

    internal static void MapMainRowStiffGeometry(
        Row row,
        string b_stiff,
        string h_stiff,
        string lws_stiff,
        string tp_stiff_map,
        string tr1_stiff_map,
        string tr2_stiff_map)
    {
        row.B_Stiff = NumericParser.ParseDouble(b_stiff);
        row.H_Stiff = NumericParser.ParseDouble(h_stiff);
        row.Lws_Stiff = NumericParser.ParseDouble(lws_stiff);
        row.Tp_Stiff = NumericParser.ParseDouble(tp_stiff_map);
        row.Tr1_Stiff = NumericParser.ParseDouble(tr1_stiff_map);
        row.Tr2_Stiff = NumericParser.ParseDouble(tr2_stiff_map);
    }

    internal static void MapMainRowBase(
        Row row,
        string f_base,
        string lws_base,
        string lp_base,
        string ls_base,
        string tws_base,
        string d_ws_base,
        string d_p_base,
        string h_base,
        string s_base,
        string b_base,
        string t_base,
        string xh_base,
        string nh_base_var1,
        string nh_base_var2)
    {
        row.F_base = NumericParser.ParseDouble(f_base);
        row.Lws_base = NumericParser.ParseDouble(lws_base);
        row.Lp_base = NumericParser.ParseDouble(lp_base);
        row.Ls_base = NumericParser.ParseDouble(ls_base);
        row.Tws_base = NumericParser.ParseDouble(tws_base);
        row.D_ws_base = NumericParser.ParseDouble(d_ws_base);
        row.D_p_base = NumericParser.ParseDouble(d_p_base);
        row.H_base = NumericParser.ParseDouble(h_base);
        row.S_base = NumericParser.ParseDouble(s_base);
        row.B_base = NumericParser.ParseDouble(b_base);
        row.T_base = NumericParser.ParseDouble(t_base);
        row.Xh_base = NumericParser.ParseDouble(xh_base);
        row.Nh_base_var1 = NumericParser.ParseDouble(nh_base_var1);
        row.Nh_base_var2 = NumericParser.ParseDouble(nh_base_var2);
    }

    internal static void MapMainRowAnchor(
        Row row,
        string anchor_var_1,
        string anchor_var_2,
        string anchor_var_3,
        string anchor_var_4)
    {
        row.Anchor_var_1 = anchor_var_1;
        row.Anchor_var_2 = anchor_var_2;
        row.Anchor_var_3 = anchor_var_3;
        row.Anchor_var_4 = anchor_var_4;
    }

    internal static void MapMainRowShearKey(
        Row row,
        string lp_shearKey,
        string ls_shearKey)
    {
        row.Lp_ShearKey = NumericParser.ParseDouble(lp_shearKey);
        row.Ls_ShearKey = NumericParser.ParseDouble(ls_shearKey);
    }

    internal static void MapMainRowBrace(
        Row row,
        string e2_mode_brace,
        string e3_mode_brace,
        string n1_mode_brace,
        string n2_mode_brace)
    {
        row.E2_Brace = NumericParser.ParseDouble(e2_mode_brace);
        row.E3_Brace = NumericParser.ParseDouble(e3_mode_brace);
        row.N1_Brace = NumericParser.ParseDouble(n1_mode_brace);
        row.N2_Brace = NumericParser.ParseDouble(n2_mode_brace);
    }

    internal static void MapProfileBeam(
        Row row,
        string profile,
        string gostProfile,
        string h,
        string b,
        string s,
        string t)
    {
        row.ProfileBeam = profile ?? "";
        row.GostBeams = gostProfile ?? "";
        row.Beam_H = NumericParser.ParseDouble(h);
        row.Beam_B = NumericParser.ParseDouble(b);
        row.Beam_s = NumericParser.ParseDouble(s);
        row.Beam_t = NumericParser.ParseDouble(t);
    }


    internal static void MapProfileBrace(
        Row row,
        string profile,
        string gostProfile,
        string h,
        string b,
        string s,
        string t)
    {
        row.ProfileBrace = profile ?? "";
        row.GostBrace = gostProfile ?? "";
        row.Brace_H = NumericParser.ParseDouble(h);
        row.Brace_B = NumericParser.ParseDouble(b);
        row.Brace_s = NumericParser.ParseDouble(s);
        row.Brace_t = NumericParser.ParseDouble(t);
    }

    internal static void MapProfileRigel(
        Row row,
        string profile,
        string gostProfile,
        string h,
        string b,
        string s,
        string t)
    {
        row.ProfileRigel = profile ?? "";

        row.GostRunThrough = gostProfile ?? "";
        row.Rigel_H = NumericParser.ParseDouble(h);
        row.Rigel_B = NumericParser.ParseDouble(b);
        row.Rigel_s = NumericParser.ParseDouble(s);
        row.Rigel_t = NumericParser.ParseDouble(t);
    }

    internal static void MapProfileRunThrough(
        Row row,
        string profile, 
        string gostProfile,
        string h,
        string b,
        string s,
        string t)
    {
        row.ProfileRunThrough = profile ?? "";
        row.GostRunThrough = gostProfile ?? "";
        row.RunThrough_H = NumericParser.ParseDouble(h);
        row.RunThrough_B = NumericParser.ParseDouble(b);
        row.RunThrough_s = NumericParser.ParseDouble(s);
        row.RunThrough_t = NumericParser.ParseDouble(t);
    }
    internal static void MapProfileColumn(
        Row row,
        string profile,
        string gostProfile,
        string h,
        string b,
        string s,
        string t)
    {
        row.ProfileColumn = profile ?? "";
        row.GostColumn = gostProfile ?? "";
        row.Column_H = NumericParser.ParseDouble(h);
        row.Column_B = NumericParser.ParseDouble(b);
        row.Column_s = NumericParser.ParseDouble(s);
        row.Column_t = NumericParser.ParseDouble(t);
    }
}
