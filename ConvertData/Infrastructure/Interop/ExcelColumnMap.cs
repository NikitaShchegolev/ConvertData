using ConvertData.Enums;

namespace ConvertData.Infrastructure.Interop;

/// <summary>
/// Карта индексов колонок Excel для отображения заголовков на свойства Row.
/// Хранит индексы всех возможных колонок из входных таблиц.
/// </summary>
internal sealed class ExcelColumnMap
{
    #region Основное
    /// <summary>Индекс колонки "Name" (имя соединения).</summary>
    public int IdxName { get; set; } = -1;
    /// <summary>Индекс колонки "StructuralElement" (тип конструктивного элемента).</summary>
    public int IdxStructuralElement { get; set; } = -1;
    /// <summary>Проверяет, является ли таблица основной (содержит Name, Code, Profile).</summary>
    public bool IsMainTable => IdxName >= 0 && IdxCode >= 0 && IdxProfileBeam >= 0;

    /// <summary>Индекс колонки "CONNECTION_CODE" (код соединения).</summary>
    public int IdxCode { get; set; } = -1;

    /// <summary> Индекс колонки "TypeNode" или "ТипУзла" (тип узла соединения).</summary>
    public int IdxTypeNode { get; set; } = -1;

    /// <summary>Индекс колонки "Gost".</summary>
    public int IdxGost { get; set; } = -1;

    /// <summary>Индекс колонки "GOST_Column_Beams".</summary>
    public int IdxGostColumnAndBeams { get; set; } = -1;

    /// <summary>Индекс колонки "GostHoles".</summary>
    public int IdxGostHoles { get; set; } = -1;

    /// <summary>Индекс колонки "GostBolts".</summary>
    public int IdxGostBolts { get; set; } = -1;

    /// <summary>Индекс колонки "GostAnchore".</summary>
    public int IdxGostAnchore { get; set; } = -1;

    /// <summary>Индекс колонки "GostWeld".</summary>
    public int IdxGostWeld { get; set; } = -1;

    /// <summary>Индекс колонки "GostColumn".</summary>
    public int IdxGostProfile { get; set; } = -1;

    /// <summary>Индекс колонки "variable" (вариант расчета).</summary>
    public int IdxVariable { get; set; } = -1;
    /// <summary>Индекс для пояснений</summary>
    public int IdxExplanations { get; set; } = -1;
    /// <summary>Индекс для марки опорного столика</summary>
    public int IdxTableBrand { get; set; } = -1;

    /// <summary>Проверяет, является ли таблица таблицей профилей (содержит Profile, H, B, s, t).</summary>
    public bool IsProfileTable => IdxProfileBeam >= 0 && IdxH_beam >= 0 && IdxB_beam >= 0 && Idxs_beam >= 0 && Idxt_beam >= 0;

    public void SetProfileSectionIndices( ProfileSectionType sectionType, int profileIndex,
        int heightIndex, int widthIndex, int wallIndex, int flangeIndex)
    {
        switch (sectionType)
        {
            case ProfileSectionType.Beam:
                IdxProfileBeam = profileIndex;
                IdxH_beam = heightIndex;
                IdxB_beam = widthIndex;
                Idxs_beam = wallIndex;
                Idxt_beam = flangeIndex;
                break;
            case ProfileSectionType.Column:
                IdxProfileColumn = profileIndex;
                IdxH_column = heightIndex;
                IdxB_column = widthIndex;
                Idxs_column = wallIndex;
                Idxt_column = flangeIndex;
                break;
            case ProfileSectionType.Brace:
                IdxProfileBrace = profileIndex;
                IdxH_brace = heightIndex;
                IdxB_brace = widthIndex;
                Idxs_brace = wallIndex;
                Idxt_brace = flangeIndex;
                break;
            case ProfileSectionType.Rigel:
                IdxProfileRigel = profileIndex;
                IdxH_rigel = heightIndex;
                IdxB_rigel = widthIndex;
                Idxs_rigel = wallIndex;
                Idxt_rigel = flangeIndex;
                break;
            case ProfileSectionType.RunThrougth:
                IdxProfileRunThrough = profileIndex;
                IdxH_runThrough = heightIndex;
                IdxB_runThrough = widthIndex;
                Idxs_runThrough = wallIndex;
                Idxt_runThrough = flangeIndex;
                break;
        }
    }
    #endregion
    #region Балка
    /// <summary>Индекс колонки "ProfileBeam" или "Профиль" (профиль балки).</summary>
    public int IdxProfileBeam { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_H" (высота балки).</summary>
    public int IdxH_beam { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_B" (ширина полки балки).</summary>
    public int IdxB_beam { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_s" (толщина стенки балки).</summary>
    public int Idxs_beam { get; set; } = -1;
    /// <summary>Индекс колонки "Beam_t" (толщина полки балки).</summary>
    public int Idxt_beam { get; set; } = -1;
    #endregion
    #region Колонна

    /// <summary>Индекс колонки "ProfileColumn" (профиль колонны).</summary>
    public int IdxProfileColumn { get; set; } = -1;
    /// <summary>Индекс колонки "Column_H" (высота колонны).</summary>
    public int IdxH_column { get; set; } = -1;
    /// <summary>Индекс колонки "Column_B" (ширина полки колонны).</summary>
    public int IdxB_column { get; set; } = -1;
    /// <summary>Индекс колонки "Column_s" (толщина стенки колонны).</summary>
    public int Idxs_column { get; set; } = -1;
    /// <summary>Индекс колонки "Column_t" (толщина полки колонны).</summary>
    public int Idxt_column { get; set; } = -1;
    #endregion
    #region Связи
    /// <summary>Индекс колонки "ProfileBrace" (профиль связи).</summary>
    public int IdxProfileBrace { get; set; } = -1;
    public int IdxH_brace { get; set; } = -1;
    public int IdxB_brace { get; set; } = -1;
    public int Idxs_brace { get; set; } = -1;
    public int Idxt_brace { get; set; } = -1;
    /// <summary>Растояние болта до края фасонки</summary>
    public int Idx_a_brace { get; set; } = -1;
    /// <summary>Растояние болта до края фасонки</summary>
    public int Idx_e2_brace { get; set; } = -1;
    /// <summary>Растояние от ребра до ряда болтов</summary>
    public int Idx_e3_brace { get; set; } = -1;
    /// <summary>Колличество болтов в 1 ряду</summary>
    public int Idx_n1_brace { get; set; } = -1;
    /// <summary>Колличество болтов в 2 ряду</summary>
    public int Idx_n2_brace { get; set; } = -1;
    public int Idx_Lb_brace { get; set; } = -1;
    #endregion
    #region Ригель
    /// <summary>Индекс колонки "ProfileRigel" (профиль ригеля).</summary>
    public int IdxProfileRigel { get; set; } = -1;
    public int IdxH_rigel { get; set; } = -1;
    public int IdxB_rigel { get; set; } = -1;
    public int Idxs_rigel { get; set; } = -1;
    public int Idxt_rigel { get; set; } = -1;
    #endregion
    #region Прогон    
    /// <summary>Индекс колонки "ProfileRunThrough" (профиль прогона).</summary>
    public int IdxProfileRunThrough { get; set; } = -1;
    public int IdxH_runThrough { get; set; } = -1;
    public int IdxB_runThrough { get; set; } = -1;
    public int Idxs_runThrough { get; set; } = -1;
    public int Idxt_runThrough { get; set; } = -1;
    #endregion
    #region Пластины
    public int IdxLb_plate { get; set; } = -1;
    public int IdxB_plate { get; set; } = -1;
    public int IdxH_plate { get; set; } = -1;
    public int IdxTp_plate { get; set; } = -1;
    public int IdxTws_plate { get; set; } = -1;
    public int IdxLws_plate { get; set; } = -1;
    public int IdxLst_plate { get; set; } = -1;
    public int IdxTr1_plate { get; set; } = -1;
    public int IdxTr2_plate { get; set; } = -1;
    #endregion
    #region Базы
    public int IdxH_base { get; set; } = -1;
    public int IdxB_base { get; set; } = -1;
    public int IdxS_base { get; set; } = -1;
    public int IdxT_base { get; set; } = -1;
    #endregion
    #region Ребра жесткости
    public int IdxB_stiff { get; set; } = -1;
    public int IdxH_stiff { get; set; } = -1;
    public int Idxtp_stiff { get; set; } = -1;
    public int IdxLws_stiff { get; set; } = -1;
    public int Idxtr1_stiff { get; set; } = -1;
    public int Idxtr2_stiff { get; set; } = -1;
    public int IdxTg_Stiff { get; set; } = -1;
    public int IdxLg_Stiff { get; set; } = -1;
    public int IdxTf_Stiff { get; set; } = -1;
    public int IdxLh_Stiff { get; set; } = -1;
    public int IdxHh_Stiff { get; set; } = -1;
    #endregion
    #region Фланец
    public int IdxTp_Flange { get; set; } = -1;
    public int IdxB_Flange { get; set; } = -1;
    public int IdxH_Flange { get; set; } = -1;
    public int IdxLb_Flange { get; set; } = -1;
    #endregion
    #region Внутренние усилия
    /// <summary> Усилие отрыва для баз </summary>
    public int IdF_base { get; set; } = -1;
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
    /// <summary>Индекс колонки "My_compression" (изгибающий момент по Y).</summary>
    public int IdxMy_compression { get; set; } = -1;
    /// <summary>Индекс колонки "My_tension" (изгибающий момент по Y).</summary>
    public int IdxMy_tension { get; set; } = -1;
    /// <summary>Индекс колонки "Mneg" (обратный момент).</summary>
    public int IdxMneg { get; set; } = -1;
    /// <summary>Индекс колонки "Mz" (изгибающий момент по Z).</summary>
    public int IdxMz { get; set; } = -1;
    /// <summary>Индекс колонки "Mz_compression" (изгибающий момент по Z).</summary>
    public int IdxMz_compression { get; set; } = -1;
    /// <summary>Индекс колонки "Mz_tension" (изгибающий момент по Z).</summary>
    public int IdxMz_tension { get; set; } = -1;
    /// <summary>Индекс колонки "Mx" (изгибающий момент по X).</summary>
    public int IdxMx { get; set; } = -1;
    /// <summary>Индекс колонки "Mw" (крутящий момент Mw).</summary>
    public int IdxMw { get; set; } = -1;
    #endregion
    #region Жесткость
    /// <summary>Индекс колонки "Sj" (жесткость Sj).</summary>
    public int IdxSj { get; set; } = -1;
    /// <summary>Индекс колонки "Sjo" (жесткость Sjo).</summary>
    public int IdxSjo { get; set; } = -1;
    #endregion
    #region Поправочные коэффициенты
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
    #endregion
    #region Геометрия анкеров
    /// <summary> Длина стороны шайбы под анкер </summary>
    public int IdLws_base { get; set; } = -1;
    /// <summary> Ширина колодца под упор </summary>
    public int IdLp_base { get; set; } = -1;
    /// <summary> Ширина противосдвигового упора в плоскости наибольшей жесткости</summary>
    public int IdLs_base { get; set; } = -1;
    /// <summary> Длина противосдвигового упора в плоскости наибольшей жесткости</summary>
    public int IdXh_base { get; set; } = -1;
    /// <summary> Толщина шайбы под анкер </summary>
    public int IdTws_base { get; set; } = -1;
    /// <summary> Диаметр отверстия в шайбе под анкер </summary>
    public int IdD_ws_base { get; set; } = -1;
    /// <summary> Диаметр отверстия под анкер </summary>
    public int IdD_p_base { get; set; } = -1;

    /// <summary> Расстояние между монтажными отверстиями </summary>
    public int IdK_fws_base { get; set; } = -1;
    #endregion
    #region Количество отверстий под анкера

    /// <summary> Количество отверстий для базы под анкера варианта 1</summary>
    public int IdNh_base_var1 { get; set; } = -1;
    /// <summary> Количество отверстий для базы под анкера варианта 2</summary>
    public int IdNh_base_var2 { get; set; } = -1;
    #endregion
    #region Тип принимаемого анкера
    /// <summary> Наименование соединения вариант 1</summary>
    public int IdAnchor_var_1 { get; set; } = -1;
    /// <summary> Наименование соединения вариант 2</summary>
    public int IdAnchor_var_2 { get; set; } = -1;
    /// <summary> Наименование соединения вариант 3</summary>
    public int IdAnchor_var_3 { get; set; } = -1;
    /// <summary> Наименование соединения вариант 4</summary>
    public int IdAnchor_var_4 { get; set; } = -1;
    #endregion
    #region Противосдвиговой упор/ShearKey
    public int IdxLp_shearKey { get; set; } = -1;
    public int IdxLs_shearKey { get; set; } = -1;
    
    #endregion
}
