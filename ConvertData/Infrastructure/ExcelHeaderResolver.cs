using System.Collections.Generic;
using System.Linq;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

/// <summary>
/// Карта индексов колонок Excel для отображения заголовков на свойства Row.
/// Хранит индексы всех возможных колонок из входных таблиц.
/// </summary>
internal sealed class ExcelColumnMap
{
    /// <summary>Индекс колонки "Beam_H" (высота балки).</summary>
    public int IdxH { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_B" (ширина полки балки).</summary>
    public int IdxB { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_s" (толщина стенки балки).</summary>
    public int Idxs { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_t" (толщина полки балки).</summary>
    public int Idxt { get; set; } = -1;
    /// <summary>Индекс колонки "Name" (имя соединения).</summary>
    public int IdxName { get; set; } = -1;
    /// <summary>Индекс колонки "CONNECTION_CODE" (код соединения).</summary>
    public int IdxCode { get; set; } = -1;
    /// <summary> Индекс колонки "TypeNode" или "ТипУзла" (тип узла соединения).</summary>
    public int IdxTypeNode { get; set; } = -1;
    /// <summary>Индекс колонки "ProfileBeam" или "Профиль" (профиль балки).</summary>
    public int IdxProfile { get; set; } = -1;
    /// <summary>Индекс колонки "ProfileColumn" (профиль колонны).</summary>
    public int IdxProfileColumn { get; set; } = -1;
    /// <summary>Индекс колонки "Nt" (усилие растяжения).</summary>
    public int IdxNt { get; set; } = -1;
    /// <summary>Индекс колонки "Qy" (поперечная сила по Y).</summary>
    public int IdxQy { get; set; } = -1;
    /// <summary>Индекс колонки "Qz" (поперечная сила по Z).</summary>
    public int IdxQz { get; set; } = -1;
    /// <summary>Индекс колонки "T" (крутящий момент).</summary>
    public int IdxT { get; set; } = -1;
    /// <summary>Индекс колонки "Nc" (усилие сжатия).</summary>
    public int IdxNc { get; set; } = -1;
    /// <summary>Индекс колонки "N" (усилие растяжения/сжатия).</summary>
    public int IdxN { get; set; } = -1;
    /// <summary>Индекс колонки "My" (изгибающий момент по Y).</summary>
    public int IdxMy { get; set; } = -1;
    /// <summary>Индекс колонки "variable" (вариант расчета).</summary>
    public int IdxVariable { get; set; } = -1;
    /// <summary>Индекс колонки "Sj" (жесткость Sj).</summary>
    public int IdxSj { get; set; } = -1;
    /// <summary>Индекс колонки "Sjo" (жесткость Sjo).</summary>
    public int IdxSjo { get; set; } = -1;
    /// <summary>Индекс колонки "Mneg" (обратный момент).</summary>
    public int IdxMneg { get; set; } = -1;
    /// <summary>Индекс колонки "Mz" (изгибающий момент по Z).</summary>
    public int IdxMz { get; set; } = -1;
    /// <summary>Индекс колонки "Mx" (изгибающий момент по X).</summary>
    public int IdxMx { get; set; } = -1;
    /// <summary>Индекс колонки "Mw" (крутящий момент Mw).</summary>
    public int IdxMw { get; set; } = -1;
    /// <summary>Индекс колонки "α" или "Alpha" (коэффициент альфа).</summary>
    public int IdxAlpha { get; set; } = -1;
    /// <summary>Индекс колонки "β" или "Beta" (коэффициент бета).</summary>
    public int IdxBeta { get; set; } = -1;
    /// <summary>Индекс колонки "γ" или "Gamma" (коэффициент гамма).</summary>
    public int IdxGamma { get; set; } = -1;
    /// <summary>Индекс колонки "δ" или "Delta" (коэффициент дельта).</summary>
    public int IdxDelta { get; set; } = -1;
    /// <summary>Индекс колонки "ε" или "Epsilon" (коэффициент эпсилон).</summary>
    public int IdxEpsilon { get; set; } = -1;
    /// <summary>Индекс колонки "λ" или "Lambda" (коэффициент лямбда).</summary>
    public int IdxLambda { get; set; } = -1;
    /// <summary>Индекс для пояснений</summary>
    public int IdxExplanations { get; set; } = -1;
    /// <summary>Проверяет, является ли таблица основной (содержит Name, Code, Profile).</summary>
    public bool IsMainTable => IdxName >= 0 && IdxCode >= 0 && IdxProfile >= 0;
    /// <summary>Проверяет, является ли таблица таблицей профилей (содержит Profile, H, B, s, t).</summary>
    public bool IsProfileTable => IdxProfile >= 0 && IdxH >= 0 && IdxB >= 0 && Idxs >= 0 && Idxt >= 0;


    #region Анкера
    /// <summary> Усилие отрыва </summary>
    public int IdF_base { get; set; } = -1;
    /// <summary> Длина стороны шайбы под анкер </summary>
    public int IdLws_base { get; set; } = -1;
    /// <summary> Ширина колодца под упор </summary>
    public int IdLp_base { get; set; } = -1;
    /// <summary> Ширина противосдвигового упора в плоскости наибольшей жесткости</summary>
    public int IdLs_base { get; set; } = -1;
    /// <summary> Толщина шайбы под анкер </summary>
    public int IdTws_base { get; set; } = -1;
    /// <summary> Диаметр отверстия в шайбе под анкер </summary>
    public int IdD_ws_base { get; set; } = -1;
    /// <summary> Диаметр отверстия под анкер </summary>
    public int IdD_p_base { get; set; } = -1;
    /// <summary> Расстояние между монтажными отверстиями </summary>
    public int IdXh_base { get; set; } = -1;
    /// <summary> Катет сварного шва крепления базы </summary>
    public int IdK_fws_base { get; set; } = -1;
    /// <summary> Количество отверстий для базы варианта 1</summary>
    public int IdNh_base_var1 { get; set; } = -1;
    /// <summary> Количество отверстий для базы варианта 2</summary>
    public int IdNh_base_var2 { get; set; } = -1;
    /// <summary> Наимернование соединения вариант 1</summary>
    public int IdAnchor_var_1 { get; set; } = -1;
    /// <summary> Наимернование соединения вариант 2</summary>
    public int IdAnchor_var_2 { get; set; } = -1;
    /// <summary> Наимернование соединения вариант 3</summary>
    public int IdAnchor_var_3 { get; set; } = -1;
    /// <summary> Наимернование соединения вариант 4</summary>
    public int IdAnchor_var_4 { get; set; } = -1;
    #endregion
}

/// <summary>
/// Разрешает заголовки колонок Excel в карту индексов для отображения на свойства Row.
/// </summary>
internal static class ExcelHeaderResolver
{
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

        var map = new ExcelColumnMap
        {
            IdxH = HeaderUtils.IndexOfHeaderAny(header, ["Beam_H"]),
            IdxB = HeaderUtils.IndexOfHeaderAny(header, ["Beam_B"]),
            Idxs = HeaderUtils.IndexOfHeaderAny(header, ["Beam_s"]),
            Idxt = HeaderUtils.IndexOfHeaderAny(header, ["Beam_t"]),
            IdxName = HeaderUtils.IndexOfHeader(header, "Name"),
            IdxCode = HeaderUtils.IndexOfHeaderAny(header, ["CONNECTION_CODE", "Connection_Code", "Code", "Код"]),
            IdxTypeNode = HeaderUtils.IndexOfHeaderAny(header, ["TypeNode", "Тип узла", "ТипУзла", "Вид узла"]),            
            IdxExplanations = HeaderUtils.IndexOfHeaderAny(header, [
                "Explanations", "Explanation",
                "Объяснения", "Объяснение",
                "Пояснения", "Пояснение",
                "Дополнения", "Дополнение",
                "Примечания", "Примечание",
                "Описание", "Описания",
                "Комментарий", "Комментарии",
                "Прим."
            ]),            
            IdxProfile = idxProfile,
            IdxProfileColumn = HeaderUtils.IndexOfHeaderAny(header, ["ProfileColumn", "Profile_Column", "ПрофильКолонны"]),
            IdxNt = HeaderUtils.IndexOfHeader(header, "Nt"),
            IdxQy = HeaderUtils.IndexOfHeaderAny(header, ["Qy"]),
            IdxQz = HeaderUtils.IndexOfHeaderAny(header, ["Qz"]),
            IdxT = HeaderUtils.IndexOfHeader(header, "T"),
            IdxNc = HeaderUtils.IndexOfHeader(header, "Nc"),
            IdxN = HeaderUtils.IndexOfHeader(header, "N"),
            IdxMy = HeaderUtils.IndexOfHeaderAny(header, ["My"]),
            IdxVariable = HeaderUtils.IndexOfHeaderAny(header, ["variable", "Variable"]),
            IdxSj = HeaderUtils.IndexOfHeader(header, "Sj"),
            IdxSjo = HeaderUtils.IndexOfHeader(header, "Sjo"),
            IdxMneg = HeaderUtils.IndexOfHeader(header, "Mneg"),
            IdxMz = HeaderUtils.IndexOfHeaderAny(header, ["Mz"]),
            IdxMx = HeaderUtils.IndexOfHeader(header, "Mx"),
            IdxMw = HeaderUtils.IndexOfHeader(header, "Mw"),
            IdF_base = HeaderUtils.IndexOfHeaderAny(header, ["F_base", "Fbase", "F_base_anchor"]),
            IdLws_base = HeaderUtils.IndexOfHeaderAny(header, ["Lws_base", "Lws", "L_ws"]),
            IdLp_base = HeaderUtils.IndexOfHeaderAny(header, ["Lp_base", "Lp", "Lpbase"]),
            IdLs_base = HeaderUtils.IndexOfHeaderAny(header, ["Ls_base", "Ls", "Lsbase"]),
            IdTws_base = HeaderUtils.IndexOfHeaderAny(header, ["Tws_base", "tws", "Tws", "tws_base"]),
            IdD_ws_base = HeaderUtils.IndexOfHeaderAny(header, ["D_ws_base", "Dws", "D_ws", "d_ws_base"]),
            IdD_p_base = HeaderUtils.IndexOfHeaderAny(header, ["D_p_base", "Dp", "D_p", "d_p_base"]),
            IdXh_base = HeaderUtils.IndexOfHeaderAny(header, ["Xh_base", "xh", "Xh", "xh_base"]),
            IdK_fws_base = HeaderUtils.IndexOfHeaderAny(header, ["K_fws_base", "kfws", "K_fws", "kfws_base"]),
            IdNh_base_var1 = HeaderUtils.IndexOfHeaderAny(header, ["Nh_base_var1", "Nh_1_2", "Nh1", "Nh_var1"]),
            IdNh_base_var2 = HeaderUtils.IndexOfHeaderAny(header, ["Nh_base_var2", "Nh_3_4", "Nh2", "Nh_var2"]),
            IdAnchor_var_1 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_1", "Анкер1", "Anchor1"]),
            IdAnchor_var_2 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_2", "Анкер2", "Anchor2"]),
            IdAnchor_var_3 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_3", "Анкер3", "Anchor3"]),
            IdAnchor_var_4 = HeaderUtils.IndexOfHeaderAny(header, ["Anchor_var_4", "Анкер4", "Anchor4"]),
        };

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

    /// <summary>
    /// Применяет логику определения колонок профиля по позициям,
    /// если таблица не распознана как основная или профильная.
    /// Предполагает, что H, B, s, t идут сразу после колонки Profile.
    /// </summary>
    /// <param name="map">Карта индексов колонок.</param>
    /// <param name="header">Список заголовков.</param>
    public static void ApplyProfileFallback(ExcelColumnMap map, List<string> header)
    {
        if (map.IsMainTable || map.IsProfileTable)
            return;

        if (map.IdxProfile >= 0)
        {
            if (map.IdxH < 0) map.IdxH = map.IdxProfile + 1;
            if (map.IdxB < 0) map.IdxB = map.IdxProfile + 2;
            if (map.Idxs < 0) map.Idxs = map.IdxProfile + 3;
            if (map.Idxt < 0) map.Idxt = map.IdxProfile + 4;
        }
        else
        {
            map.IdxProfile = 0;
            map.IdxH = 1;
            map.IdxB = 2;
            map.Idxs = 3;
            map.Idxt = 4;
        }
    }
}
