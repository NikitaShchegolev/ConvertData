using ConvertData.Domain;
using ConvertData.Enums;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

internal static partial class EpplusRowPropertyMaps
{
    private static readonly ProfileSectionType[] GeometryProfileSections =
    [
        ProfileSectionType.Beam,
        ProfileSectionType.Column,
        ProfileSectionType.Brace,
        ProfileSectionType.Rigel,
        ProfileSectionType.RunThrougth
    ];

    internal static bool TryGetSheetMap(string sheetName, out Dictionary<string, Action<Row, string>> propertyMap)
    {
        if (string.Equals(sheetName, "geometry", StringComparison.OrdinalIgnoreCase))
        {
            propertyMap = GeometryColumnMap;
            return true;
        }

        if (string.Equals(sheetName, "bolts", StringComparison.OrdinalIgnoreCase))
        {
            propertyMap = BoltsColumnMap;
            return true;
        }

        if (string.Equals(sheetName, "holes", StringComparison.OrdinalIgnoreCase))
        {
            propertyMap = HolesColumnMap;
            return true;
        }

        if (string.Equals(sheetName, "weld", StringComparison.OrdinalIgnoreCase))
        {
            propertyMap = WeldColumnMap;
            return true;
        }

        propertyMap = null!;
        return false;
    }    
    internal static readonly string[] KeyColumnHeaders =  ["CONNECTION_CODE"];
    #region Работа с листами в excel
    internal static readonly Dictionary<string, Action<Row, string>> GeometryColumnMap = BuildGeometryColumnMap();
    internal static readonly Dictionary<string, Action<Row, string>> WeldColumnMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["GostWeld"] = (r, v) => r.GostWeld = v,
        ["kf1"] = (r, v) => r.kf1 = v,
        ["kf2"] = (r, v) => r.kf2 = v,
        ["kf3"] = (r, v) => r.kf3 = v,
        ["kf4"] = (r, v) => r.kf4 = v,
        ["kf5"] = (r, v) => r.kf5 = v,
        ["kf6"] = (r, v) => r.kf6 = v,
        ["kf7"] = (r, v) => r.kf7 = v,
        ["kf8"] = (r, v) => r.kf8 = v,
        ["kf9"] = (r, v) => r.kf9 = v,
        ["kf10"] = (r, v) => r.kf10 = v,
        ["k_fws"] = (r, v) => r.K_fws_base = v
    };
    internal static readonly Dictionary<string, Action<Row, string>> ShearKeyColumnMap = ShearKeyMap();
    internal static readonly Dictionary<string, Action<Row, string>> AnchorColumnMap = AnchorMap();
    internal static readonly Dictionary<string, Action<Row, string>> BoltsColumnMap = MergeMaps(MergeMaps(BoltsMap(), ShearKeyColumnMap), AnchorColumnMap);
    internal static readonly Dictionary<string, Action<Row, string>> HolesColumnMap = MergeMaps(HolesMap(), AnchorColumnMap); 
    #endregion
    private static Dictionary<string, Action<Row, string>> BuildGeometryColumnMap()
    {
        var map = new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var sectionType in GeometryProfileSections) { AddGeometryProfileSection(map, sectionType); }
            
        AddGeometryBraceBolt(map);
        AddGeometryPlate(map);
        AddGeometryFlange(map);
        AddGeometryBase(map);
        AddGeometryStiff(map);
        return map;
    }
    private static void AddGeometryProfileSection(Dictionary<string, Action<Row, string>> map, ProfileSectionType sectionType)
    {
        switch (sectionType)
        {
            case ProfileSectionType.Beam:
                AddGeometryBeam(map);
                break;
            case ProfileSectionType.Column:
                AddGeometryColumns(map);
                break;
            case ProfileSectionType.Brace:
                AddGeometryBrace(map);
                break;
            case ProfileSectionType.Rigel:
                AddGeometryRigel(map);
                break;
            case ProfileSectionType.RunThrougth:
                AddGeometryRunThrough(map);
                break;
        }
    }    
    private static void AddGeometryBraceBolt(Dictionary<string, Action<Row, string>> map)
    {
        AddNumericColumn(map, (r, v) => r.Lb_Brace = v, "Lb_brace");
        AddNumericColumn(map, (r, v) => r.Tp_Brace = v, "Tp_brace");
        AddNumericColumn(map, (r, v) => r.A_Brace = v, "a_brace");
        AddNumericColumn(map, (r, v) => r.E2_Brace = v, "e2_brace");
        AddNumericColumn(map, (r, v) => r.E3_Brace = v, "e3_brace");
        AddNumericColumn(map, (r, v) => r.N1_Brace = v, "n1_brace");
        AddNumericColumn(map, (r, v) => r.N2_Brace = v, "n2_brace");
    }
    private static void AddGeometryBase(Dictionary<string, Action<Row, string>> map)
    {
        map["F_base"] = (r, v) => r.F_base = NumericParser.ParseDouble(v);
        map["H_base"] = (r, v) => r.H_base = NumericParser.ParseDouble(v);
        map["B_base"] = (r, v) => r.B_base = NumericParser.ParseDouble(v);
        map["S_base"] = (r, v) => r.S_base = NumericParser.ParseDouble(v);
        map["Lp_base"] = (r, v) => r.Lp_base = NumericParser.ParseDouble(v);
        map["Ls_base"] = (r, v) => r.Ls_base = NumericParser.ParseDouble(v);
        map["Lws_base"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v);
        map["T_base"] = (r, v) => r.T_base = NumericParser.ParseDouble(v);
        map["Tws_base"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v);
        map["Dws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v);
        map["D_ws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v);
        map["Dp_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v);
        map["D_p_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v);
        map["K_fws_base"] = (r, v) => r.K_fws_base = v;
        map["k_fws_base"] = (r, v) => r.K_fws_base = v;
        map["Nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v);
        map["Nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v);
    }
    private static void AddGeometryPlate(Dictionary<string, Action<Row, string>> map)
    {
        map["H_plate"] = (r, v) => r.H_Plate = NumericParser.ParseDouble(v);
        map["B_plate"] = (r, v) => r.B_Plate = NumericParser.ParseDouble(v);
        map["Lb_plate"] = (r, v) => r.Lb_Plate = NumericParser.ParseDouble(v);
        map["Lws_plate"] = (r, v) => r.Lws_Plate = NumericParser.ParseDouble(v);
        map["Tp_plate"] = (r, v) => r.Tp_Plate = NumericParser.ParseDouble(v);
        map["Tr1_plate"] = (r, v) => r.Tr1_Plate = NumericParser.ParseDouble(v);
        map["Tr2_plate"] = (r, v) => r.Tr2_Plate = NumericParser.ParseDouble(v);
    }
    private static void AddGeometryFlange(Dictionary<string, Action<Row, string>> map)
    {
        map["Lb_flange"] = (r, v) => r.Flange_Lb = NumericParser.ParseDouble(v);
        map["H_flange"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Flange_H = value;
            if (r.H_Plate == 0)
                r.H_Plate = value;
        };
        map["B_flange"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Flange_B = value;
            if (r.B_Plate == 0)
                r.B_Plate = value;
        };
        map["Tp_flange"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Flange_t = value;
            if (r.Tp_Plate == 0)
                r.Tp_Plate = value;
        };
    }
    private static void AddGeometryStiff(Dictionary<string, Action<Row, string>> map)
    {
        map["B_stiff"] = (r, v) => r.B_Stiff = NumericParser.ParseDouble(v);
        map["H_stiff"] = (r, v) =>
        {
            if (r.H_Stiff == 0)
                r.H_Stiff = NumericParser.ParseDouble(v);
        };
        map["Hh_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Hh_Stiff = value;
            if (r.Hh_Stiff == 0)
                r.Hh_Stiff = value;
        };
        map["Lh_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Lh_Stiff = value;
            if (r.Lh_Stiff == 0)
                r.Lh_Stiff = value;
        };
        map["Lg_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Lg_Stiff = value;
            if (r.Lg_Stiff == 0)
                r.Lg_Stiff = value;
        };
        map["Lws_stiff"] = (r, v) => r.Lws_Stiff = NumericParser.ParseDouble(v);
        map["Tp_stiff"] = (r, v) => r.Tp_Stiff = NumericParser.ParseDouble(v);
        map["Tr1_stiff"] = (r, v) => r.Tr1_Stiff = NumericParser.ParseDouble(v);
        map["Tr2_stiff"] = (r, v) => r.Tr2_Stiff = NumericParser.ParseDouble(v);
        map["Tg_stiff"] = (r, v) => r.Tg_Stiff = NumericParser.ParseDouble(v);
        map["Tf_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Tf_Stiff = value;
            if (r.Tf_Stiff == 0)
                r.Tf_Stiff = value;
        };
        map["Twp_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Twp_Stiff = value;
            if (r.Twp_Stiff == 0)
                r.Twp_Stiff = value;
        };
        map["Tbp_stiff"] = (r, v) =>
        {
            var value = NumericParser.ParseDouble(v);
            r.Tbp_Stiff = value;
            if (r.Tbp_Stiff == 0)
                r.Tbp_Stiff = value;
        };
    }
    /// <summary>
    /// Метод отвечает за добавление в карту свойств геометрических характеристик для заданного типа профиля (балка, колонна, связь, ригель, прогон).
    /// </summary>
    /// <param name="map"></param>
    /// <param name="section"></param>
    private static void AddGeometrySection(Dictionary<string, Action<Row, string>> map, GeometrySectionDefinition section)
    {
        AddTextColumn(map, section.ProfileSetter, section.ProfileHeaders);
        AddTextColumn(map, section.GostSetter, section.GostHeaders);
        AddNumericColumn(map, section.HSetter, section.GetHeaders("H"));
        AddNumericColumn(map, section.BSetter, section.GetHeaders("B"));
        AddNumericColumn(map, section.SSetter, section.GetHeaders("s"));
        AddNumericColumn(map, section.TSetter, section.GetHeaders("t"));
        AddNumericColumn(map, section.ASetter, section.GetHeaders("A"));
        AddNumericColumn(map, section.PSetter, section.GetHeaders("P"));
        AddNumericColumn(map, section.IzSetter, section.GetHeaders("Iz"));
        AddNumericColumn(map, section.IySetter, section.GetHeaders("Iy"));
        AddNumericColumn(map, section.IxSetter, section.GetHeaders("Ix"));
        AddNumericColumn(map, section.WzSetter, section.GetHeaders("Wz"));
        AddNumericColumn(map, section.WySetter, section.GetHeaders("Wy"));
        AddNumericColumn(map, section.WxSetter, section.GetHeaders("Wx"));
        AddNumericColumn(map, section.SzSetter, section.GetHeaders("Sz"));
        AddNumericColumn(map, section.SySetter, section.GetHeaders("Sy"));
        AddNumericColumn(map, section.izSetter, section.GetHeaders("iz"));
        AddNumericColumn(map, section.iySetter, section.GetHeaders("iy"));
        AddNumericColumn(map, section.xoSetter, section.GetHeaders("xo"));
        AddNumericColumn(map, section.yoSetter, section.GetHeaders("yo"));
    }
    /// <summary>
    /// Метод возвращает определение геометрической секции
    /// (балка, колонна, связь, ригель, прогон) с соответствующими сеттерами и
    /// заголовками для заполнения свойств Row. Это позволяет избежать 
    /// дублирования кода при добавлении различных типов профилей в карту свойств.
    /// </summary>
    /// <param name="sectionType"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    private static GeometrySectionDefinition GetGeometrySectionDefinition(ProfileSectionType sectionType)
    {
        return sectionType switch
        {
            ProfileSectionType.Beam => new GeometrySectionDefinition
            {
                FieldPrefix = "Beam",
                ProfileSetter = (r, v) => r.ProfileBeam = v,
                ProfileHeaders = ["ProfileBeam", "ProfileBeams"],
                GostSetter = (r, v) => r.GostBeams = v,
                GostHeaders = ["GostBeam"],
                HSetter = (r, v) => r.Beam_H = v,
                BSetter = (r, v) => r.Beam_B = v,
                SSetter = (r, v) => r.Beam_s = v,
                TSetter = (r, v) => r.Beam_t = v,
                ASetter = (r, v) => r.Beam_A = v,
                PSetter = (r, v) => r.Beam_P = v,
                IzSetter = (r, v) => r.Beam_Iz = v,
                IySetter = (r, v) => r.Beam_Iy = v,
                IxSetter = (r, v) => r.Beam_Ix = v,
                WzSetter = (r, v) => r.Beam_Wz = v,
                WySetter = (r, v) => r.Beam_Wy = v,
                WxSetter = (r, v) => r.Beam_Wx = v,
                SzSetter = (r, v) => r.Beam_Sz = v,
                SySetter = (r, v) => r.Beam_Sy = v,
                izSetter = (r, v) => r.Beam_iz = v,
                iySetter = (r, v) => r.Beam_iy = v,
                xoSetter = (r, v) => r.Beam_xo = v,
                yoSetter = (r, v) => r.Beam_yo = v                
            },
            ProfileSectionType.Column => new GeometrySectionDefinition
            {
                FieldPrefix = "Column",
                ProfileSetter = (r, v) => r.ProfileColumn = v,
                ProfileHeaders = ["ProfileColumn", "ProfileColumns"],
                GostSetter = (r, v) => r.GostColumn = v,
                GostHeaders = ["GostColumn"],
                HSetter = (r, v) => r.Column_H = v,
                BSetter = (r, v) => r.Column_B = v,
                SSetter = (r, v) => r.Column_s = v,
                TSetter = (r, v) => r.Column_t = v,
                ASetter = (r, v) => r.Column_A = v,
                PSetter = (r, v) => r.Column_P = v,
                IzSetter = (r, v) => r.Column_Iz = v,
                IySetter = (r, v) => r.Column_Iy = v,
                IxSetter = (r, v) => r.Column_Ix = v,
                WzSetter = (r, v) => r.Column_Wz = v,
                WySetter = (r, v) => r.Column_Wy = v,
                WxSetter = (r, v) => r.Column_Wx = v,
                SzSetter = (r, v) => r.Column_Sz = v,
                SySetter = (r, v) => r.Column_Sy = v,
                izSetter = (r, v) => r.Column_iz = v,
                iySetter = (r, v) => r.Column_iy = v,
                xoSetter = (r, v) => r.Column_xo = v,
                yoSetter = (r, v) => r.Column_yo = v
            },
            ProfileSectionType.Brace => new GeometrySectionDefinition
            {
                FieldPrefix = "Brace",
                ProfileSetter = (r, v) => r.ProfileBrace = v,
                ProfileHeaders = ["ProfileBrace"],
                GostSetter = (r, v) => r.GostBrace = v,
                GostHeaders = ["GostBrace"],
                HSetter = (r, v) => r.Brace_H = v,
                BSetter = (r, v) => r.Brace_B = v,
                SSetter = (r, v) => r.Brace_s = v,
                TSetter = (r, v) => r.Brace_t = v,
                ASetter = (r, v) => r.Brace_A = v,
                PSetter = (r, v) => r.Brace_P = v,
                IzSetter = (r, v) => r.Brace_Iz = v,
                IySetter = (r, v) => r.Brace_Iy = v,
                IxSetter = (r, v) => r.Brace_Ix = v,
                WzSetter = (r, v) => r.Brace_Wz = v,
                WySetter = (r, v) => r.Brace_Wy = v,
                WxSetter = (r, v) => r.Brace_Wx = v,
                SzSetter = (r, v) => r.Brace_Sz = v,
                SySetter = (r, v) => r.Brace_Sy = v,
                izSetter = (r, v) => r.Brace_iz = v,
                iySetter = (r, v) => r.Brace_iy = v,
                xoSetter = (r, v) => r.Brace_xo = v,
                yoSetter = (r, v) => r.Brace_yo = v
            },
            ProfileSectionType.Rigel => new GeometrySectionDefinition
            {
                FieldPrefix = "Rigel",
                ProfileSetter = (r, v) => r.ProfileRigel = v,
                ProfileHeaders = ["ProfileRigel"],
                GostSetter = (r, v) => r.GostRigel = v,
                GostHeaders = ["GostRigel"],
                HSetter = (r, v) => r.Rigel_H = v,
                BSetter = (r, v) => r.Rigel_B = v,
                SSetter = (r, v) => r.Rigel_s = v,
                TSetter = (r, v) => r.Rigel_t = v,
                ASetter = (r, v) => r.Rigel_A = v,
                PSetter = (r, v) => r.Rigel_P = v,
                IzSetter = (r, v) => r.Rigel_Iz = v,
                IySetter = (r, v) => r.Rigel_Iy = v,
                IxSetter = (r, v) => r.Rigel_Ix = v,
                WzSetter = (r, v) => r.Rigel_Wz = v,
                WySetter = (r, v) => r.Rigel_Wy = v,
                WxSetter = (r, v) => r.Rigel_Wx = v,
                SzSetter = (r, v) => r.Rigel_Sz = v,
                SySetter = (r, v) => r.Rigel_Sy = v,
                izSetter = (r, v) => r.Rigel_iz = v,
                iySetter = (r, v) => r.Rigel_iy = v,
                xoSetter = (r, v) => r.Rigel_xo = v,
                yoSetter = (r, v) => r.Rigel_yo = v
                
            },
            ProfileSectionType.RunThrougth => new GeometrySectionDefinition
            {
                FieldPrefix = "RunThrougth",
                ProfileSetter = (r, v) => r.ProfileRunThrough = v,
                ProfileHeaders = ["ProfileRunThrough"],
                GostSetter = (r, v) => r.GostRunThrougth = v,
                GostHeaders = ["GostRunThrougth"],
                HSetter = (r, v) => r.RunThrougth_H = v,
                BSetter = (r, v) => r.RunThrougth_B = v,
                SSetter = (r, v) => r.RunThrougth_s = v,
                TSetter = (r, v) => r.RunThrougth_t = v,
                ASetter = (r, v) => r.RunThrougth_A = v,
                PSetter = (r, v) => r.RunThrougth_P = v,
                IzSetter = (r, v) => r.RunThrougth_Iz = v,
                IySetter = (r, v) => r.RunThrougth_Iy = v,
                IxSetter = (r, v) => r.RunThrougth_Ix = v,
                WzSetter = (r, v) => r.RunThrougth_Wz = v,
                WySetter = (r, v) => r.RunThrougth_Wy = v,
                WxSetter = (r, v) => r.RunThrougth_Wx = v,
                SzSetter = (r, v) => r.RunThrougth_Sz = v,
                SySetter = (r, v) => r.RunThrougth_Sy = v,
                izSetter = (r, v) => r.RunThrougth_iz = v,
                iySetter = (r, v) => r.RunThrougth_iy = v,
                xoSetter = (r, v) => r.RunThrougth_xo = v,
                yoSetter = (r, v) => r.RunThrougth_yo = v
            },
            _ => throw new ArgumentOutOfRangeException(nameof(sectionType), sectionType, null)
        };
    }
   /// <summary>
   /// Заполнение геометрических характеристик для балки
   /// </summary>
   /// <param name="map"></param>
    private static void AddGeometryBeam(Dictionary<string, Action<Row, string>> map)
    {
        AddGeometrySection(map, GetGeometrySectionDefinition(ProfileSectionType.Beam));
    }
    /// <summary>
    /// Заполнение геометрических характеристик для колонны
    /// </summary>
    /// <param name="map"></param>
    private static void AddGeometryColumns(Dictionary<string, Action<Row, string>> map)
    {
        AddGeometrySection(map, GetGeometrySectionDefinition(ProfileSectionType.Column));
    }
    /// <summary>
    /// Заполнение геометрических характеристик для связи
    /// </summary>
    /// <param name="map"></param>
    private static void AddGeometryBrace(Dictionary<string, Action<Row, string>> map)
    {
        AddGeometrySection(map, GetGeometrySectionDefinition(ProfileSectionType.Brace));
    }
    /// <summary>
    /// Заполнение геометрических характеристик для ригеля
    /// </summary>
    /// <param name="map"></param>
    private static void AddGeometryRigel(Dictionary<string, Action<Row, string>> map)
    {
        AddGeometrySection(map, GetGeometrySectionDefinition(ProfileSectionType.Rigel));
    }
    /// <summary>
    /// Заполнение геометрических характеристик для прогона
    /// </summary>
    /// <param name="map"></param>
    private static void AddGeometryRunThrough(Dictionary<string, Action<Row, string>> map)
    {
        AddGeometrySection(map, GetGeometrySectionDefinition(ProfileSectionType.RunThrougth));
    }
       
    /// <summary>
    /// Заполнение болтов
    /// </summary>
    /// <returns></returns>
    private static Dictionary<string, Action<Row, string>> BoltsMap()
    {
        return new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Option"] = (r, v) => r.OptionBolts = NumericParser.ParseInt(v),
            ["TypeNode"] = (r, v) => r.TypeNode = v,
            ["GostBolts"] = (r, v) => r.GostBolts = v,
            ["F"] = (r, v) => r.F = NumericParser.ParseInt(v),
            ["N_rows"] = (r, v) => r.N_Rows = NumericParser.ParseInt(v),
            ["Nb"] = (r, v) => r.Bolts_Nb = NumericParser.ParseInt(v),
            ["e1"] = (r, v) => r.e1 = NumericParser.ParseInt(v),
            ["p1"] = (r, v) => r.p1 = NumericParser.ParseInt(v),
            ["p2"] = (r, v) => r.p2 = NumericParser.ParseInt(v),
            ["p3"] = (r, v) => r.p3 = NumericParser.ParseInt(v),
            ["p4"] = (r, v) => r.p4 = NumericParser.ParseInt(v),
            ["p5"] = (r, v) => r.p5 = NumericParser.ParseInt(v),
            ["p6"] = (r, v) => r.p6 = NumericParser.ParseInt(v),
            ["p7"] = (r, v) => r.p7 = NumericParser.ParseInt(v),
            ["p8"] = (r, v) => r.p8 = NumericParser.ParseInt(v),
            ["p9"] = (r, v) => r.p9 = NumericParser.ParseInt(v),
            ["p10"] = (r, v) => r.p10 = NumericParser.ParseInt(v),
            ["d1"] = (r, v) =>
            {
                EpplusWorksheetHelpers.EnsureBolts(r, 1);
                r.CoordinatesBolts[0].X = NumericParser.ParseInt(v);
            },
            ["d2"] = (r, v) =>
            {
                EpplusWorksheetHelpers.EnsureBolts(r, 2);
                r.CoordinatesBolts[1].X = NumericParser.ParseInt(v);
                if (r.N_Rows < 2)
                    r.N_Rows = 2;
            },
            ["a_brace"] = (r, v) => r.A_Brace = NumericParser.ParseDouble(v),
            ["e2_brace"] = (r, v) => r.E2_Brace = NumericParser.ParseDouble(v),
            ["e3_brace"] = (r, v) => r.E3_Brace = NumericParser.ParseDouble(v),
            ["n1_brace"] = (r, v) => r.N1_Brace = NumericParser.ParseDouble(v),
            ["n2_brace"] = (r, v) => r.N2_Brace = NumericParser.ParseDouble(v)
        };
    }

    private static Dictionary<string, Action<Row, string>> ShearKeyMap()
    {
        return new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Lp_shearKey"] = (r, v) => r.Lp_ShearKey = NumericParser.ParseDouble(v),
            ["Ls_shearKey"] = (r, v) => r.Ls_ShearKey = NumericParser.ParseDouble(v)
        };
    }

    private static Dictionary<string, Action<Row, string>> AnchorMap()
    {
        return new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["GostAnchore"] = (r, v) => r.GostAnchore = v,
            ["Nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v),
            ["Nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v),
            ["Anchor_var_1"] = (r, v) => r.Anchor_var_1 = v,
            ["Anchor_var_2"] = (r, v) => r.Anchor_var_2 = v,
            ["Anchor_var_3"] = (r, v) => r.Anchor_var_3 = v,
            ["Anchor_var_4"] = (r, v) => r.Anchor_var_4 = v
        };
    }

    private static Dictionary<string, Action<Row, string>> MergeMaps(
        Dictionary<string, Action<Row, string>> left,
        Dictionary<string, Action<Row, string>> right)
    {
        var map = new Dictionary<string, Action<Row, string>>(left, StringComparer.OrdinalIgnoreCase);

        foreach (var pair in right)
            map[pair.Key] = pair.Value;

        return map;
    }

    private static Dictionary<string, Action<Row, string>> HolesMap()
    {
        return new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Option"] = (r, v) => r.OptionHoles = NumericParser.ParseInt(v),
            ["F_holes"] = (r, v) => r.F_holes = NumericParser.ParseInt(v),
            ["Dws_holes"] = (r, v) => r.Dws_holes = NumericParser.ParseDouble(v),
            ["Dp_holes"] = (r, v) => r.Dp_holes = NumericParser.ParseDouble(v),
            ["xh"] = (r, v) => r.xh_holes = NumericParser.ParseDouble(v),
            ["Nh_holes_1_4"] = (r, v) => r.Nh_Holes_1_4 = NumericParser.ParseInt(v),
            ["Nh_holes_5_8"] = (r, v) => r.Nh_Holes_5_8 = NumericParser.ParseInt(v)
        };
    }
    private static void AddTextColumn(Dictionary<string, Action<Row, string>> map, Action<Row, string> setter,params string[] headers)
    {
        foreach (var header in headers)
            map[header] = setter;
    }
    private static void AddNumericColumn(Dictionary<string, Action<Row, string>> map, Action<Row, double> setter,params string[] headers)
    {
        foreach (var header in headers)
            map[header] = (r, v) => setter(r, NumericParser.ParseDouble(v));
    }
}
