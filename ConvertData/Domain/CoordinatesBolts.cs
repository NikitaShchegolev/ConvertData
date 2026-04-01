namespace ConvertData.Domain;

/// <summary>
/// Класс для расчета координат болтов 
/// относительно расстояния между рядами 
/// координат расстояния по ширина и высоте
/// </summary>
internal class CoordinatesBolts
{
    /// <summary>
    /// Координата болта в направлении x
    /// </summary>
    public int X { get; set; }
    /// <summary>
    /// Координата болта в направлении y
    /// </summary>
    public int Y { get; set; }
    /// <summary>
    /// Координата болта в направлении z
    /// </summary>
    public int Z { get; set; }
    /// <summary>
    /// Конструктор для создания координат болта.
    /// </summary>
    /// <param name="x">Координата по оси X.</param>
    /// <param name="y">Координата по оси Y.</param>
    /// <param name="z">Координата по оси Z.</param>
    public CoordinatesBolts(int x, int y, int z)
    {
        X = x;
        Y = y;
        Z = z;
    }

}
