using System.Collections.Generic;
using System.Linq;
using ConvertData.Enums;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure.Interop;

/// <summary>
/// Разрешает заголовки колонок Excel в карту индексов для отображения на свойства Row.
/// </summary>
internal static class ExcelHeaderResolver
{
    private sealed class ProfileSectionHeaderInfo
    {
        public ProfileSectionHeaderInfo(
            string[] profileHeaders,
            string[] heightHeaders,
            string[] widthHeaders,
            string[] wallThicknessHeaders,
            string[] flangeThicknessHeaders)
        {
            ProfileHeaders = profileHeaders;
            HeightHeaders = heightHeaders;
            WidthHeaders = widthHeaders;
            WallThicknessHeaders = wallThicknessHeaders;
            FlangeThicknessHeaders = flangeThicknessHeaders;
        }

        public string[] ProfileHeaders { get; }
        public string[] HeightHeaders { get; }
        public string[] WidthHeaders { get; }
        public string[] WallThicknessHeaders { get; }
        public string[] FlangeThicknessHeaders { get; }
    }

    private static readonly ProfileSectionType[] ProfileSections =
    [
        ProfileSectionType.Beam,
        ProfileSectionType.Column,
        ProfileSectionType.Brace,
        ProfileSectionType.Rigel,
        ProfileSectionType.RunThrougth
    ];

    private static readonly Dictionary<ProfileSectionType, ProfileSectionHeaderInfo> ProfileSectionHeaders = new()
    {
        [ProfileSectionType.Beam] = new(["ProfileBeams", "ProfileBeam"], ["Beam_H"], ["Beam_B"], ["Beam_s"], ["Beam_t"]),
        [ProfileSectionType.Column] = new(["ProfileColumns", "ProfileColumn"], ["Column_H"], ["Column_B"], ["Column_s"], ["Column_t"]),
        [ProfileSectionType.Brace] = new(["ProfileBrace", "ProfileBraces"], ["Brace_H"], ["Brace_B"], ["Brace_s"], ["Brace_t"]),
        [ProfileSectionType.Rigel] = new(["ProfileRigel"], ["Rigel_H"], ["Rigel_B"], ["Rigel_s"], ["Rigel_t"]),
        [ProfileSectionType.RunThrougth] = new(["ProfileRunThrough", "ProfileRunThrougth", "ProfileRunTrought"], ["RunThrougth_H"], ["RunThrougth_B"], ["RunThrougth_s"], ["RunThrougth_t"])
    };

    /// <summary>
    /// Переопределение имени колонки профиля из аргументов командной строки (--profile-column).
    /// </summary>
    public static string? ProfileColumnOverride { get; set; }

    /// <summary>
    /// Разрешает список заголовков в карту индексов колонок.
    /// </summary>
    /// <param name="header">Список нормализованных заголовков из Excel.</param>
    /// <returns>Карта индексов колонок.</returns>
    public static ExcelColumnMap Resolve(List<string> header)
    {
        int idxProfile;
        if (!string.IsNullOrWhiteSpace(ProfileColumnOverride))
        {
            idxProfile = HeaderUtils.IndexOfHeader(header, ProfileColumnOverride);
            if (idxProfile < 0)
                idxProfile = HeaderUtils.IndexOfHeaderAny(header, ["ProfileBeam", "Профиль"]);
        }
        else
        {
            idxProfile = HeaderUtils.IndexOfHeaderAny(header, ["ProfileBeam", "Профиль"]);
        }
        //IdxProfileBeam = idxProfile, - не знаю для чего!
        var map = new ExcelColumnMap
        {
            //Общие данные об узле
            IdxName = HeaderUtils.IndexOfHeader(header, "Name"),
            IdxCode = HeaderUtils.IndexOfHeaderAny(header, ["CONNECTION_CODE", "Connection_Code", "Code", "Код"]),
            IdxTypeNode = HeaderUtils.IndexOfHeaderAny(header, ["TypeNode", "Тип узла", "ТипУзла", "Вид узла"]),
            IdxGostColumnAndBeams = HeaderUtils.IndexOfHeaderAny(header, ["GostBeams", "GostColumnAndBeams", "GOST_Column_Beams", "Gost_Column_Beams", "GOST Column Beams"]),
            IdxGostBolts = HeaderUtils.IndexOfHeaderAny(header, ["GostBolts", "GOST_bolts"]),
            IdxGostAnchore = HeaderUtils.IndexOfHeaderAny(header, ["GostAnchore", "GOST_anchor", "GOST_anchors"]),
            IdxGostWeld = HeaderUtils.IndexOfHeaderAny(header, ["GostWeld", "GOST_weld"]),
            IdxGostProfile = HeaderUtils.IndexOfHeaderAny(header, ["GostColumn", "GOST_Profile", "Gost_Profile", "GOST Profile"]),
            IdxExplanations = HeaderUtils.IndexOfHeaderAny(header, ["Explanations", "Explanation"]),
            IdxTableBrand = HeaderUtils.IndexOfHeaderAny(header, ["Марка опорного столика", "Маркаопорногостолика", "Марка"]),


            //Внутренние усилия
            IdF_base = HeaderUtils.IndexOfHeaderAny(header, ["F_base", "Fbase", "F_base_anchor"]),
            IdxNt = HeaderUtils.IndexOfHeader(header, "Nt"),
            IdxQy = HeaderUtils.IndexOfHeaderAny(header, ["Qy"]),
            IdxQz = HeaderUtils.IndexOfHeaderAny(header, ["Qz"]),
            IdxT = HeaderUtils.IndexOfHeader(header, "T"),
            IdxNc = HeaderUtils.IndexOfHeader(header, "Nc"),
            IdxN = HeaderUtils.IndexOfHeader(header, "N"),
            IdxMy = HeaderUtils.IndexOfHeaderAny(header, ["My"]),
            IdxMy_compression = HeaderUtils.IndexOfHeaderAny(header, ["My_compression", "My_compresion", "My_ compresion", "My compression"]),
            IdxMy_tension = HeaderUtils.IndexOfHeaderAny(header, ["My_tension", "My_ tension", "My tension"]),
            IdxMz = HeaderUtils.IndexOfHeaderAny(header, ["Mz"]),
            IdxMz_compression = HeaderUtils.IndexOfHeaderAny(header, ["Mz_compression", "Mz_compresion", "Mz_ compresion", "Mz compression"]),
            IdxMz_tension = HeaderUtils.IndexOfHeaderAny(header, ["Mz_tension", "Mz_ tension", "Mz tension"]),
            IdxMneg = HeaderUtils.IndexOfHeader(header, "Mneg"),
            IdxMx = HeaderUtils.IndexOfHeader(header, "Mx"),
            IdxMw = HeaderUtils.IndexOfHeader(header, "Mw"),
            IdxVariable = HeaderUtils.IndexOfHeaderAny(header, ["variable", "Variable"]),
            //Жесткость
            IdxSj = HeaderUtils.IndexOfHeader(header, "Sj"),
            IdxSjo = HeaderUtils.IndexOfHeader(header, "Sjo"),


            IdLws_base = HeaderUtils.IndexOfHeaderAny(header, ["Lws_base", "Lws", "L_ws"]),
            IdLp_base = HeaderUtils.IndexOfHeaderAny(header, ["Lp_base"]),
            IdLs_base = HeaderUtils.IndexOfHeaderAny(header, ["Ls_base"]),
            IdTws_base = HeaderUtils.IndexOfHeaderAny(header, ["Tws_base", "tws", "Tws", "tws_base"]),
            IdD_ws_base = HeaderUtils.IndexOfHeaderAny(header, ["D_ws_base", "Dws", "D_ws", "d_ws_base"]),
            IdD_p_base = HeaderUtils.IndexOfHeaderAny(header, ["D_p_base", "Dp", "D_p", "d_p_base"]),
            IdXh_base = HeaderUtils.IndexOfHeaderAny(header, ["Xh_base", "xh", "Xh", "xh_base"]),
            IdK_fws_base = HeaderUtils.IndexOfHeaderAny(header, ["K_fws_base"]),
            IdNh_base_var1 = HeaderUtils.IndexOfHeaderAny(header, ["Nh_base_var1"]),
            IdNh_base_var2 = HeaderUtils.IndexOfHeaderAny(header, ["Nh_base_var2"]),
            IdxH_base = HeaderUtils.IndexOfHeaderAny(header, ["H_base"]),
            IdxB_base = HeaderUtils.IndexOfHeaderAny(header, ["B_base"]),
            IdxS_base = HeaderUtils.IndexOfHeaderAny(header, ["S_base"]),
            IdxT_base = HeaderUtils.IndexOfHeaderAny(header, ["T_base"]),

            IdxLb_plate = HeaderUtils.IndexOfHeaderAny(header, ["Lb_plate"]),
            IdxB_plate = HeaderUtils.IndexOfHeaderAny(header, ["B_plate", "Plate_B"]),
            IdxH_plate = HeaderUtils.IndexOfHeaderAny(header, ["H_plate", "Plate_H"]),
            IdxLws_plate = HeaderUtils.IndexOfHeaderAny(header, ["Lws_plate", "Plate_Lws"]),
            IdxTp_plate = HeaderUtils.IndexOfHeaderAny(header, ["tp_plate", "Tp_plate", "Plate_t", "Plate_tp"]),
            IdxTr1_plate = HeaderUtils.IndexOfHeaderAny(header, ["tr1_plate", "Tr1_plate", "Plate_tr1"]),
            IdxTr2_plate = HeaderUtils.IndexOfHeaderAny(header, ["tr2_plate", "Tr2_plate", "Plate_tr2"]),

            IdxB_stiff = HeaderUtils.IndexOfHeaderAny(header, ["B_stiff", "Stiff_B"]),
            IdxH_stiff = HeaderUtils.IndexOfHeaderAny(header, ["H_stiff", "Stiff_H"]),
            IdxLws_stiff = HeaderUtils.IndexOfHeaderAny(header, ["Lws_stiff", "Stiff_Lws"]),
            Idxtp_stiff = HeaderUtils.IndexOfHeaderAny(header, ["tp_stiff", "Tp_stiff", "Stiff_tp"]),
            Idxtr1_stiff = HeaderUtils.IndexOfHeaderAny(header, ["tr1_stiff", "Tr1_stiff", "Stiff_tr1"]),
            Idxtr2_stiff = HeaderUtils.IndexOfHeaderAny(header, ["tr2_stiff", "Tr2_stiff", "Stiff_tr2"]),
            IdxTg_Stiff = HeaderUtils.IndexOfHeaderAny(header, ["Tg_stiff"]),
            IdxLg_Stiff = HeaderUtils.IndexOfHeaderAny(header, ["Lg_stiff"]),
            IdxTf_Stiff = HeaderUtils.IndexOfHeaderAny(header, ["Tf_stiff"]),
            IdxLh_Stiff = HeaderUtils.IndexOfHeaderAny(header, ["Lh_stiff"]),
            IdxHh_Stiff = HeaderUtils.IndexOfHeaderAny(header, ["Hh_stiff"]),

            IdxTp_Flange = HeaderUtils.IndexOfHeaderAny(header, ["Tp_flange"]),
            IdxB_Flange = HeaderUtils.IndexOfHeaderAny(header, ["B_flange"]),
            IdxH_Flange = HeaderUtils.IndexOfHeaderAny(header, ["H_flange"]),
            IdxLb_Flange = HeaderUtils.IndexOfHeaderAny(header, ["Lb_flange"]),


            IdAnchor_var_1 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_1"]),
            IdAnchor_var_2 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_2"]),
            IdAnchor_var_3 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_3"]),
            IdAnchor_var_4 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_4"]),

            IdxLp_shearKey = HeaderUtils.IndexOfHeaderAny(header, ["Lp_shearKey"]),
            IdxLs_shearKey = HeaderUtils.IndexOfHeaderAny(header, ["Ls_shearKey"]),

            Idx_a_brace = HeaderUtils.IndexOfHeaderAny(header, ["a_brace"]),
            Idx_e2_brace = HeaderUtils.IndexOfHeaderAny(header, ["e2_brace"]),
            Idx_e3_brace = HeaderUtils.IndexOfHeaderAny(header, ["e3_brace"]),
            Idx_n1_brace = HeaderUtils.IndexOfHeaderAny(header, ["n1_brace"]),
            Idx_n2_brace = HeaderUtils.IndexOfHeaderAny(header, ["n2_brace"]),
            Idx_Lb_brace = HeaderUtils.IndexOfHeaderAny(header, ["Lb_brace"])
        };

        foreach (var sectionType in ProfileSections)
            ResolveProfileSectionHeaders(header, map, sectionType);

        map.IdxAlpha = HeaderUtils.IndexOfHeader(header, "α");
        if (map.IdxAlpha < 0) map.IdxAlpha = HeaderUtils.IndexOfHeader(header, "Alpha");
        map.IdxBeta = HeaderUtils.IndexOfHeader(header, "β");
        if (map.IdxBeta < 0) map.IdxBeta = HeaderUtils.IndexOfHeader(header, "Beta");
        map.IdxGamma = HeaderUtils.IndexOfHeader(header, "γ");
        if (map.IdxGamma < 0) map.IdxGamma = HeaderUtils.IndexOfHeader(header, "Gamma");
        map.IdxDelta = HeaderUtils.IndexOfHeader(header, "δ");
        if (map.IdxDelta < 0) map.IdxDelta = HeaderUtils.IndexOfHeader(header, "Delta");
        map.IdxEpsilon = HeaderUtils.IndexOfHeader(header, "ε");
        if (map.IdxEpsilon < 0) map.IdxEpsilon = HeaderUtils.IndexOfHeader(header, "Epsilon");
        map.IdxLambda = HeaderUtils.IndexOfHeader(header, "λ");
        if (map.IdxLambda < 0) map.IdxLambda = HeaderUtils.IndexOfHeader(header, "Lambda");

        ResolveGreekFallback(header, map);

        return map;
    }

    private static void ResolveProfileSectionHeaders(List<string> header, ExcelColumnMap map, ProfileSectionType sectionType)
    {
        if (!ProfileSectionHeaders.TryGetValue(sectionType, out var headerInfo))
            return;

        var profileIndex = HeaderUtils.IndexOfHeaderAny(header, headerInfo.ProfileHeaders);
        var heightIndex = HeaderUtils.IndexOfHeaderAny(header, headerInfo.HeightHeaders);
        var widthIndex = HeaderUtils.IndexOfHeaderAny(header, headerInfo.WidthHeaders);
        var wallIndex = HeaderUtils.IndexOfHeaderAny(header, headerInfo.WallThicknessHeaders);
        var flangeIndex = HeaderUtils.IndexOfHeaderAny(header, headerInfo.FlangeThicknessHeaders);

        map.SetProfileSectionIndices(sectionType, profileIndex, heightIndex, widthIndex, wallIndex, flangeIndex);
    }

    /// <summary>
    /// Пытается определить индексы греческих коэффициентов (α, β, γ, δ, ε, λ),
    /// если они не были найдены по заголовкам. Использует позиционную логику или "?" заголовки.
    /// </summary>
    /// <param name="header">Список заголовков.</param>
    /// <param name="map">Карта индексов колонок.</param>
    private static void ResolveGreekFallback(List<string> header, ExcelColumnMap map)
    {
        if (map.IdxMz < 0)
            return;
        if (map.IdxAlpha >= 0 && map.IdxBeta >= 0 && map.IdxGamma >= 0
            && map.IdxDelta >= 0 && map.IdxEpsilon >= 0 && map.IdxLambda >= 0)
            return;

        var qMarks = header
            .Select((h, i) => new { h, i })
            .Where(x => x.h == "?")
            .Select(x => x.i)
            .ToList();

        int baseIdx = map.IdxMz + 1;
        if (baseIdx < header.Count && header.Count - baseIdx >= 6)
        {
            if (map.IdxAlpha < 0) map.IdxAlpha = baseIdx + 0;
            if (map.IdxBeta < 0) map.IdxBeta = baseIdx + 1;
            if (map.IdxGamma < 0) map.IdxGamma = baseIdx + 2;
            if (map.IdxDelta < 0) map.IdxDelta = baseIdx + 3;
            if (map.IdxEpsilon < 0) map.IdxEpsilon = baseIdx + 4;
            if (map.IdxLambda < 0) map.IdxLambda = baseIdx + 5;
        }
        else if (qMarks.Count >= 6)
        {
            if (map.IdxAlpha < 0) map.IdxAlpha = qMarks[0];
            if (map.IdxBeta < 0) map.IdxBeta = qMarks[1];
            if (map.IdxGamma < 0) map.IdxGamma = qMarks[2];
            if (map.IdxDelta < 0) map.IdxDelta = qMarks[3];
            if (map.IdxEpsilon < 0) map.IdxEpsilon = qMarks[4];
            if (map.IdxLambda < 0) map.IdxLambda = qMarks[5];
        }
    }

}
