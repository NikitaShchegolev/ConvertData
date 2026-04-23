using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

using System.Reflection.PortableExecutable;

using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ConvertData.Infrastructure;

/// <summary>
/// Отображает строковые данные из Excel в объекты Row.
/// Преобразует текстовые значения в соответствующие числовые типы.
/// </summary>
internal static class RowMapper
{
    /// <summary>
    /// Отображает строку из основной таблицы Excel в объект Row со всеми параметрами соединения.
    /// </summary>
    /// <param name="name">Имя соединения.</param>
    /// <param name="code">Код соединения (CONNECTION_CODE).</param>
    /// <param name="typeNode">Тип узла.</param>
    /// <param name="explanations">Пояснения.</param>
    /// <param name="profileBeam">Профиль балки.</param>
    /// <param name="profileColumn">Профиль колонны.</param>
    /// <param name="h">Высота сечения балки.</param>
    /// <param name="b">Ширина полки балки.</param>
    /// <param name="s">Толщина стенки балки.</param>
    /// <param name="tGeom">Толщина полки балки.</param>
    /// <param name="nt">Усилие растяжения.</param>
    /// <param name="q">Поперечная сила Qy.</param>
    /// <param name="qz">Поперечная сила Qz.</param>
    /// <param name="t">Крутящий момент T.</param>
    /// <param name="nc">Усилие сжатия.</param>
    /// <param name="n">Усилие растяжения/сжатия.</param>
    /// <param name="my">Изгибающий момент My.</param>
    /// <param name="variable">Вариант расчета.</param>
    /// <param name="sj">Жесткость Sj.</param>
    /// <param name="sjo">Жесткость Sjo.</param>
    /// <param name="mneg">Обратный момент.</param>
    /// <param name="mz">Изгибающий момент Mz.</param>
    /// <param name="mx">Изгибающий момент Mx.</param>
    /// <param name="mw">Крутящий момент Mw.</param>
    /// <param name="alpha">Коэффициент α.</param>
    /// <param name="beta">Коэффициент β.</param>
    /// <param name="gamma">Коэффициент γ.</param>
    /// <param name="delta">Коэффициент δ.</param>
    /// <param name="epsilon">Коэффициент ε.</param>
    /// <param name="lambda">Коэффициент λ.</param>
    /// <param name="lws_base">Расстояние между точками крепления.</param>
    /// <param name="lp_base">Расстояние между точками крепления.</param>
    /// <param name="ls_base">Расстояние между точками крепления.</param>
    /// <param name="tws_base">Толщина пластины для базы.</param>
    /// <param name="d_ws_base"></param>
    /// <param name="d_p_base"></param>
    /// <param name="xh_base">Расстояние</param>
    /// <param name="k_fws_base">Сварка для базы.</param>
    /// <param name="nh_base_var1">Вариант.</param>
    /// <param name="nh_base_var2">Вариант</param>
    /// <param name= "anchor_var_1">Вариант анкера.</param>
    /// <param name= "anchor_var_2">Вариант анкера.</param>
    /// <param name= "anchor_var_3">Вариант анкера.</param>
    /// <param name= "anchor_var_4">Вариант анкера.</param>
    /// <returns>Объект Row с заполненными свойствами.</returns>
    public static Row MapMainRow(
        string name,
        string code,
        string typeNode,
        string gost,
        string gostColumnAndBeams,
        string gostProfile,
        string profileBeam,
        string profileColumn,
        string explanations,
        string h,
        string b,
        string s,
        string tGeom,
        string nt,
        string q,
        string qz,
        string t,
        string nc,
        string n,
        string my,
        string my_compression,
        string my_tension,
        string variable,
        string sj,
        string sjo,
        string mneg,
        string mz,
        string mz_compression,
        string mz_tension,
        string mx,
        string mw,
        string alpha,
        string beta,
        string gamma,
        string delta,
        string epsilon,
        string lambda,
        string bB_plate,
        string hH_plate,
        string lws_plate,
        string tp_plate_map,
        string tr1_plate_map,
        string tr2_plate_map,
        string b_stiff,
        string h_stiff,
        string lws_stiff,
        string tp_stiff_map,
        string tr1_stiff_map,
        string tr2_stiff_map,
        string f_base,
        string lws_base,
        string lp_base,
        string ls_base,
        string tws_base,
        string d_ws_base,
        string d_p_base,
        string xh_base,
        string nh_base_var1,
        string nh_base_var2,
        string anchor_var_1,
        string anchor_var_2,
        string anchor_var_3,
        string anchor_var_4

        )
    {
        return new Row
        {
            Name = name,
            CONNECTION_CODE = code,
            TypeNode = typeNode ?? "",
            Gost = gost ?? "",
            GostColumnAndBeams = gostColumnAndBeams ?? "",
            GostProfile = gostProfile ?? "",
            ProfileBeam = profileBeam ?? "",
            ProfileColumn = profileColumn ?? "",
            Explanations = explanations ?? "",
            variable = variable ?? "",
            Sj = NumericParser.ParseInt(sj),
            Sjo = NumericParser.ParseInt(sjo),
            Beam_H = NumericParser.ParseDouble(h),
            Beam_B = NumericParser.ParseDouble(b),
            Beam_s = NumericParser.ParseDouble(s),
            Beam_t = NumericParser.ParseDouble(tGeom),
            Nt = NumericParser.ParseInt(nt),
            Nc = NumericParser.ParseInt(nc),
            N = NumericParser.ParseInt(n),
            Qz = NumericParser.ParseInt(qz),
            Qy = NumericParser.ParseInt(q),
            My = NumericParser.ParseInt(my),
            My_compression = NumericParser.ParseInt(my_compression),
            My_tension = NumericParser.ParseInt(my_tension),
            Mz = NumericParser.ParseDouble(mz),
            Mz_compression = NumericParser.ParseDouble(mz_compression),
            Mz_tension = NumericParser.ParseDouble(mz_tension),
            Mx = NumericParser.ParseDouble(mx),
            Mw = NumericParser.ParseDouble(mw),
            T = NumericParser.ParseInt(t),
            Mneg = NumericParser.ParseDouble(mneg),
            Alpha = NumericParser.ParseDouble(alpha),
            Beta = NumericParser.ParseDouble(beta),
            Gamma = NumericParser.ParseDouble(gamma),
            Delta = NumericParser.ParseDouble(delta),
            Epsilon = NumericParser.ParseDouble(epsilon),
            Lambda = NumericParser.ParseDouble(lambda),
            B_Plate = NumericParser.ParseDouble(bB_plate),
            H_Plate = NumericParser.ParseDouble(hH_plate),
            Lws_Plate   = NumericParser.ParseDouble(lws_plate),
            Tp_Plate = NumericParser.ParseDouble(tp_plate_map),
            Tr1_Plate   = NumericParser.ParseDouble(tr1_plate_map),
            Tr2_Plate   = NumericParser.ParseDouble(tr2_plate_map),
            B_Stiff = NumericParser.ParseDouble(b_stiff),
            H_Stiff = NumericParser.ParseDouble(h_stiff),
            Lws_Stiff   = NumericParser.ParseDouble(lws_stiff),
            Tp_Stiff   = NumericParser.ParseDouble(tp_stiff_map),
            Tr1_Stiff   = NumericParser.ParseDouble(tr1_stiff_map),
            Tr2_Stiff   = NumericParser.ParseDouble(tr2_stiff_map),
            F_base = NumericParser.ParseDouble(f_base),
            Lws_base = NumericParser.ParseDouble(lws_base),
            Lp_base = NumericParser.ParseDouble(lp_base),
            Ls_base = NumericParser.ParseDouble(ls_base),
            Tws_base = NumericParser.ParseDouble(tws_base),
            D_ws_base = NumericParser.ParseDouble(d_ws_base),
            D_p_base = NumericParser.ParseDouble(d_p_base),
            Xh_base = NumericParser.ParseDouble(xh_base),
            Nh_base_var1 = NumericParser.ParseDouble(nh_base_var1),
            Nh_base_var2 = NumericParser.ParseDouble(nh_base_var2),
            Anchor_var_1 = anchor_var_1,
            Anchor_var_2 = anchor_var_2,
            Anchor_var_3 = anchor_var_3,
            Anchor_var_4 = anchor_var_4
        };
    }

    /// <summary>
    /// Отображает строку из таблицы профилей в объект Row с геометрическими параметрами балки.
    /// </summary>
    /// <param name="profile">Профиль балки.</param>
    /// <param name="h">Высота сечения балки.</param>
    /// <param name="b">Ширина полки балки.</param>
    /// <param name="s">Толщина стенки балки.</param>
    /// <param name="t">Толщина полки балки.</param>
    /// <returns>Объект Row с заполненными геометрическими свойствами балки.</returns>
    public static Row MapProfileRow(string profile, string gostProfile, string h, string b, string s, string t)
    {
        return new Row
        {
            ProfileBeam = profile,
            GostProfile = gostProfile ?? "",
            Beam_H = NumericParser.ParseDouble(h),
            Beam_B = NumericParser.ParseDouble(b),
            Beam_s = NumericParser.ParseDouble(s),
            Beam_t = NumericParser.ParseDouble(t)
        };
    }
}
