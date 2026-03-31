namespace ConvertData.Domain;

internal sealed record ProfileGeometry
{
    /// <summary>
    /// Высота сечения балки
    /// </summary>
    public double H { get; init; }
    /// <summary>
    /// Ширина пояса балки
    /// </summary>
    public double B { get; init; }
    /// <summary>
    /// Толщина стенки балки
    /// </summary>
    public double t_w { get; init; }
    /// <summary>
    /// Толщина полки балкии
    /// </summary>
    public double t_f { get; init; }
    /// <summary>
    /// Радиус скругления либо полки либо между балкой и стенкой
    /// </summary>
    public double r1 { get; init; }
    /// <summary>
    /// Радиус скругления либо полки либо между балкой и стенкой
    /// </summary>
    public double r2 { get; init; }
    /// <summary>
    /// Площадь сечения профиля балки
    /// </summary>
    public double A { get; init; }
    /// <summary>
    /// Собственный вес метра балки
    /// </summary>
    public double P { get; init; }
    /// <summary>
    /// Момент инерции балки в плоскости наименьшей жескости
    /// </summary>
    public double Iz { get; init; }
    /// <summary>
    /// Момент инерции балки в плоскости наибольшей жескости
    /// </summary>
    public double Iy { get; init; }
    /// <summary>
    /// Момент инерции вокруг оси
    /// </summary>
    public double Ix { get; init; }
    /// <summary>
    /// Момент инерции для уголка
    /// </summary>
    public double Iv { get; init; }
    /// <summary>
    /// Момент инерции для уголка
    /// </summary>
    public double Iyz { get; init; }
    /// <summary>
    /// Момен с сопротивления балки в плоскости наименьшей жескости
    /// </summary>
    public double Wz { get; init; }
    /// <summary>
    /// Момент сопротивления балки в плоскости наибольшей жескости
    /// </summary>
    public double Wy { get; init; }
    /// <summary>
    /// Момент сопротивления балки вокруг оси
    /// </summary>
    public double Wx { get; init; }
    /// <summary>
    /// Момент сопротивления для уголка
    /// </summary>
    public double Wvo { get; init; }
    /// <summary>
    /// Статический момент сечения балки в плоскости наименьшей жескости
    /// </summary>
    public double Sz { get; init; }
    /// <summary>
    /// Статический момент сечения балки в плоскости наибольшей жескости
    /// </summary>
    public double Sy { get; init; }
    /// <summary>
    /// Радиус инерции балки в плоскости наименьшей жескости
    /// </summary>
    public double iz { get; init; }
    /// <summary>
    /// Радиус инерции балки в плоскости наибольшей жескости
    /// </summary>
    public double iy { get; init; }
    /// <summary>
    /// Координаты центра тяжести сечения балки 
    /// относительно условного начала координат по оси X
    /// </summary>
    public double xo { get; init; }
    /// <summary>
    /// Координаты центра тяжести сечения балки 
    /// относительно условного начала координат по оси Y
    /// </summary>
    public double yo { get; init; }
    /// <summary>
    /// Координаты центра тяжести сечения для уголка
    /// </summary>
    public double iu { get; init; }
    /// <summary>
    /// Координаты центра тяжести сечения для уголка
    /// </summary>
    public double iv { get; init; }
}
