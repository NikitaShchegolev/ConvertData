namespace ConvertData.Domain;

internal sealed class Row
{
    /// <summary>
    /// Имя группы узлового соединения
    /// </summary>
    public string Name { get; set; } = "";
    public string TypeNode { get; set; } = "";
    /// <summary>
    /// Пояснения к узловому соединению, которые могут включать в себя описание типа соединения,
    /// </summary>
    public string Explanations { get; set; } = "";
    /// <summary>
    /// Код соединения, который определяет тип соединения между балкой и колонной. 
    /// Этот код может быть использован для определения геометрических 
    /// характеристик балки и колонны, а также для расчета жесткости и внутренних сил в соединении.
    /// </summary>
    public string CONNECTION_CODE { get; set; } = "";
    #region Вариант расчета
    /// <summary>
    /// Вариант расчета, который определяет, какие данные будут использоваться для расчета.
    /// </summary>
    public int variable { get; set; }
    #endregion


    // Geometry
    #region Beam - геометрические характеристики балки
    /// <summary>
    /// "Beam": профиль балки
    /// </summary>
    public string ProfileBeam { get; set; } = "";
    /// <summary>
    /// "Beam": Высота сечения балки
    /// </summary>
    public double Beam_H { get; set; }
    /// <summary>
    /// "Beam": Ширина полки балки
    /// </summary>
    public double Beam_B { get; set; }
    /// <summary>
    /// "Beam": толщина стенки балки
    /// </summary>
    public double Beam_s { get; set; }
    /// <summary>
    /// "Beam":  толщина полки балки
    /// </summary>
    public double Beam_t { get; set; }
    /// <summary>
    /// "Beam": Площадь сечения балки
    /// </summary>
    public double Beam_A { get; set; }
    /// <summary>
    /// "Beam": Вес метра погонного балки
    /// </summary>
    public double Beam_P { get; set; }
    /// <summary>
    /// "Beam": Момент инерции балки относительно плоскости наименьшей жескости
    /// </summary>
    public double Beam_Iz { get; set; }
    /// <summary>
    /// "Beam": Момент инерции балки относительно плоскости наибольшей жескости
    /// </summary>
    public double Beam_Iy { get; set; }
    /// <summary>
    /// "Beam": Момент инерции балки относительно центральной оси x
    /// </summary>
    public double Beam_Ix { get; set; }
    /// </summary>
    /// "Beam":  Момент сопротивления балки относительно центральной оси z
    /// </summary>
    public double Beam_Wz { get; set; }
    /// <summary>
    /// "Beam": Момент сопротивления балки относительно центральной оси y
    /// </summary>
    public double Beam_Wy { get; set; }
    /// "Beam": Момент сопротивления балки относительно центральной оси x
    /// </summary>
    public double Beam_Wx { get; set; }

    /// "Beam": Статический момент балки относительно центральной оси z
    /// </summary>
    public double Beam_Sz { get; set; }
    /// <summary>
    /// "Beam": Статический момент балки относительно центральной оси y
    /// </summary>
    public double Beam_Sy { get; set; }
    /// <summary>
    /// "Beam": Радиус инерции балки относительно центральной оси z
    /// </summary>
    public double Beam_iz { get; set; }
    /// <summary>
    /// "Beam": Радиус инерции балки относительно центральной оси y
    /// </summary>
    public double Beam_iy { get; set; }
    /// <summary>
    /// "Beam": Координата центра тяжести балки в направления x
    /// </summary>
    public double Beam_xo { get; set; }
    /// <summary>
    /// "Beam": Координата центра тяжести балки в направления y
    /// </summary>
    public double Beam_yo { get; set; }
    #endregion
    #region Column - геометрическин характеристики колонны
    /// <summary>
    /// "Column": профиль колонны
    /// </summary>
    public string ProfileColumn { get; set; } = "";
    /// <summary>
    /// "Column": Высота сечения
    /// </summary>
    public double Column_H { get; set; }
    /// <summary>
    /// "Column": Ширина полки колонны
    /// </summary>
    public double Column_B { get; set; }
    /// <summary>
    /// "Column": толщна стенки колонны
    /// </summary>
    public double Column_s { get; set; }
    /// <summary>
    /// "Column": толщна полки колонны
    /// </summary>
    public double Column_t { get; set; }
    /// <summary>
    /// "Column":
    /// </summary>
    public double Column_A { get; set; }

    /// <summary>
    /// "Column": Вес метра погонного колонны
    /// </summary>
    public double Column_P { get; set; }
    /// <summary>
    /// "Column": Момент инерции колонны относительно плоскости
    /// </summary>
    public double Column_Iz { get; set; }
    /// <summary>
    /// "Column": Момент инерции колонны относительно плоскости
    /// </summary>
    public double Column_Iy { get; set; }
    /// <summary>
    /// "Column": Момент инерции колонны относительно плоскости
    /// </summary>
    public double Column_Ix { get; set; }
    /// <summary>
    /// "Column": Момент сопротивления колонны относительно центральной оси z
    /// </summary>
    public double Column_Wz { get; set; }
    /// <summary>
    /// "Column": Момент сопротивления колонны относительно центральной оси y
    /// </summary>
    public double Column_Wy { get; set; }
    /// <summary>
    /// "Column": Момент сопротивления колонны относительно центральной оси x
    /// </summary>
    public double Column_Wx { get; set; }
    /// <summary>
    /// "Column": Статический момент колонны относительно центральной оси z
    /// </summary>
    public double Column_Sz { get; set; }
    /// <summary>
    /// "Column": Статический момент колонны относительно центральной оси y
    /// </summary>
    public double Column_Sy { get; set; }
    /// <summary>
    /// "Column":
    /// </summary>
    public double Column_iz { get; set; }
    /// <summary>
    /// "Column": Радиус инерции колонны относительно центральной оси z
    /// </summary>
    public double Column_iy { get; set; }
    /// <summary>
    /// "Column": Координата центра тяжести колонны в направления x
    /// </summary>
    public double Column_xo { get; set; }
    /// <summary>
    /// "Column": Координата центра тяжести балки в направления y
    /// </summary>
    public double Column_yo { get; set; }    
    #endregion
    #region Plate
    /// <summary>
    /// "Plate": Длина пластины
    /// </summary>
    public double Plate_H { get; set; }
    /// <summary>
    /// "Plate": Ширина пластины
    /// </summary>
    public double Plate_B { get; set; }
    /// <summary>
    /// "Plate": толщна пластины
    /// </summary>
    public double Plate_t { get; set; }
    #endregion
    #region Flange
    /// <summary>
    /// "Flange": Расстояние от верха полки колонны до края пластины для фланца, который крепится к полке колонны
    /// </summary>
    public double Flange_Lb { get; set; }
    /// <summary>
    /// "Flange": Длина пластины
    /// </summary>
    public double Flange_H { get; set; }
    /// <summary>
    /// "Flange": Ширина пластины
    /// </summary>
    public double Flange_B { get; set; }
    /// <summary>
    /// "Flange": толщна пластины
    /// </summary>
    public double Flange_t { get; set; }
    #endregion
    #region Stiff
    /// <summary>
    /// "Stiff": Расстояние от верха полки колонны до края пластины для фланца, который крепится к полке колонны
    /// </summary>
    public double Stiff_tbp { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_tg { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_tf { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_Lh { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_Hh { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_tr1 { get; set; }
    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_tr2 { get; set; }

    /// <summary>
    /// "Stiff": 
    /// </summary>
    public double Stiff_twp { get; set; }
    #endregion

    //Bolts

    #region Bolts
    /// <summary>
    /// "Bolts": Список коордниат для одного болта
    /// </summary>
    public List<CoordinatesBolts> CoordinatesBolts { get; set; } = new List<CoordinatesBolts>();
    
    /// <summary>
    /// "Bolts": Диаметр болта
    /// </summary>
    public int F { get; set; }

    /// <summary>
    /// "Bolts": Количество болтов
    /// </summary>
    public int Bolts_Nb { get; set; }

    /// <summary>
    /// "Bolts": Количество рядов болтов
    /// </summary>
    public int N_Rows { get; set; }

    /// <summary>
    /// "Bolts": Версия использования болтов
    /// </summary>
    public double OptionBolts { get; set; } = 0;
    /// <summary>
    /// "Bolts": Марка опорного столика
    /// </summary>
    public string TableBrand { get; set; } = "";
    /// <summary>
    /// "Bolts": Координата Y первого болта (расстояние от края пластины)
    /// </summary>
    public int e1 { get; set; }
    /// <summary>
    /// "Bolts": Координата X первого ряда болтов
    /// </summary>
    public int d1 { get; set; }
    /// <summary>
    /// "Bolts": Координата X второго ряда болтов
    /// </summary>
    public int d2 { get; set; }
    /// <summary>
    /// Расстояние от края пластины до 1 ряда болтов
    /// </summary>
    public double p1 { get; set; }
    /// <summary>
    /// Расстояние между 1 и 2 рядом болтов
    /// </summary>
    public double p2 { get; set; }
    /// <summary>
    /// Расстояние между 2 и 3 рядом болтов
    /// </summary>
    public double p3 { get; set; }
    /// <summary>
    /// Расстояние между 3 и 4 рядом болтов
    /// </summary>
    public double p4 { get; set; }
    /// <summary>
    /// Расстояние между 4 и 5 рядом болтов
    /// </summary>
    public double p5 { get; set; }
    /// <summary>
    /// Расстояние между 5 и 6 рядом болтов
    /// </summary>
    public double p6 { get; set; }
    /// <summary>
    /// Расстояние между 6 и 7 рядом болтов
    /// </summary>
    public double p7 { get; set; }
    /// <summary>
    /// Расстояние между 7 и 8 рядом болтов
    /// </summary>
    public double p8 { get; set; }
    /// <summary>
    /// Расстояние между 8 и 9 рядом болтов
    /// </summary>
    public double p9 { get; set; }
    /// <summary>
    /// Расстояние между 9 и 10 рядом болтов
    /// </summary>
    public double p10 { get; set; }


    #endregion


    //Welds - Минимальные катеты сварных швов

    #region Weld

    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf1 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf2 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf3 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf4 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf5 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf6 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf7 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf8 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf9 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string kf10 { get; set; } = "";
    /// <summary>
    /// Минимальный катет сварного шва
    /// </summary>
    public string K_fws_base { get; set; } = "";
    #endregion

    //Характеристики материала

    #region Stiffness - жесткость
    /// <summary>
    ///  "Stiffness" - Sj, жесткость в направлении Uy
    /// </summary>
    public int Sj { get; set; }
    /// <summary>
    /// "Stiffness" - Sjo, жесткость в направлении Uz
    /// </summary>
    public int Sjo { get; set; }
    #endregion
    #region InternalForces - внутренние силы
    /// <summary>
    /// "InternalForces" - Nt, растяжение
    /// </summary>
    public int Nt { get; set; }
    /// <summary>
    /// "InternalForces" - Nt, Сжатие
    /// </summary>
    public int Nc { get; set; }
    /// <summary>
    /// "InternalForces" - Nt, растяжение/сжатие
    /// </summary>
    public int N { get; set; }
    /// <summary>
    /// "InternalForces" - Qz, поперечная сила в направлении z
    /// </summary>
    public int Qz { get; set; }
    /// <summary>
    /// "InternalForces" - Qy, поперечная сила в направлении y
    /// </summary>
    public int Qy { get; set; }

    /// <summary>
    /// "InternalForces" - Qx, поперечная сила в направлении x
    /// </summary>
    public int Qx { get; set; }
    /// <summary>
    /// "InternalForces" - My, изгибающий момент в направлении y
    /// </summary>
    public int My { get; set; }
    /// <summary>
    /// "InternalForces" - Mneg - обратный момент
    /// </summary>
    public double Mneg { get; set; }
    /// <summary>
    /// "InternalForces" - Mz, изгибающий момент в направлении z
    /// </summary>
    public double Mz { get; set; }
    /// <summary>
    /// "InternalForces" - Mx, изгибающий момент в направлении x
    /// </summary>
    public double Mx { get; set; }
    /// <summary>
    /// "InternalForces" - Mw, крутящий момент
    /// </summary>
    public double Mw { get; set; }
    /// <summary>
    /// "InternalForces" - T, крутящий момент
    /// </summary>
    public int T { get; set; }
    #endregion
    #region Coefficients - расчетные коэффициенты
    /// <summary>
    /// "Coefficients" - α
    /// </summary>
    public double Alpha { get; set; }
    /// <summary>
    /// "Coefficients" - β
    /// </summary>
    public double Beta { get; set; }
    /// <summary>
    /// "Coefficients" - γ
    /// </summary>
    public double Gamma { get; set; }
    /// <summary>
    /// "Coefficients" - δ
    /// </summary>
    public double Delta { get; set; }
    /// <summary>
    /// "Coefficients" - ε
    /// </summary>
    public double Epsilon { get; set; }
    /// <summary>
    /// "Coefficients" - λ
    /// </summary>
    public double Lambda { get; set; }
    #endregion

    //Анкера

    #region Анкера
    /// <summary> Усилие отрыва </summary>
    public double F_base { get; set; }
    /// <summary> Длина стороны шайбы под анкер </summary>
    public double Lws_base { get; set; }
    /// <summary> Ширина колодца под упор </summary>
    public double Lp_base { get; set; }
    /// <summary> Ширина противосдвигового упора в плоскости наибольшей жесткости</summary>
    public double Ls_base { get; set; }
    /// <summary> Толщина шайбы под анкер </summary>
    public double Tws_base { get; set; }
    /// <summary> Диаметр отверстия в шайбе под анкер </summary>
    public double D_ws_base { get; set; }
    /// <summary> Диаметр отверстия под анкер </summary>
    public double D_p_base { get; set; }
    /// <summary> Расстояние между монтажными отверстиями </summary>
    public double Xh_base { get; set; }
    /// <summary> Катет сварного шва крепления базы </summary>
    /// <summary> Количество отверстий для базы варианта 1</summary>
    public double Nh_base_var1 { get; set; }
    /// <summary> Количество отверстий для базы варианта 2</summary>
    public double Nh_base_var2 { get; set; } 
    /// <summary> Наимернование соединения вариант 1</summary>
    public string Anchor_var_1 { get; set; } = "";
    /// <summary> Наимернование соединения вариант 2</summary>
    public string Anchor_var_2 { get; set; } = "";
    /// <summary> Наимернование соединения вариант 3</summary>
    public string Anchor_var_3 { get; set; } = "";
    /// <summary> Наимернование соединения вариант 4</summary>
    public string Anchor_var_4 { get; set; } = ""; 
    #endregion

}