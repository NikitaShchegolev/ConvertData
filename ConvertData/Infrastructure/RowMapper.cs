using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

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
    /// <returns>Объект Row с заполненными свойствами.</returns>
    public static Row MapMainRow(
        string name,
        string code,
        string typeNode,
        string profileBeam,
        string profileColumn,
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
        string variable,
        string sj,
        string sjo,
        string mneg,
        string mz,
        string mx,
        string mw,
        string alpha,
        string beta,
        string gamma,
        string delta,
        string epsilon,
        string lambda)
    {
        return new Row
        {
            Name = name,
            CONNECTION_CODE = code,
            TypeNode = typeNode ?? "",
            ProfileBeam = profileBeam ?? "",
            ProfileColumn = profileColumn ?? "",
            variable = NumericParser.ParseInt(variable),
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
            Mz = NumericParser.ParseDouble(mz),
            Mx = NumericParser.ParseDouble(mx),
            Mw = NumericParser.ParseDouble(mw),
            T = NumericParser.ParseInt(t),
            Mneg = NumericParser.ParseDouble(mneg),
            Alpha = NumericParser.ParseDouble(alpha),
            Beta = NumericParser.ParseDouble(beta),
            Gamma = NumericParser.ParseDouble(gamma),
            Delta = NumericParser.ParseDouble(delta),
            Epsilon = NumericParser.ParseDouble(epsilon),
            Lambda = NumericParser.ParseDouble(lambda)
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
    public static Row MapProfileRow(string profile, string h, string b, string s, string t)
    {
        return new Row
        {
            ProfileBeam = profile,
            Beam_H = NumericParser.ParseDouble(h),
            Beam_B = NumericParser.ParseDouble(b),
            Beam_s = NumericParser.ParseDouble(s),
            Beam_t = NumericParser.ParseDouble(t)
        };
    }
}
