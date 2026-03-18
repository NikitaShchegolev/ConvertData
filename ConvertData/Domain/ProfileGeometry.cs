namespace ConvertData.Domain;

internal sealed record ProfileGeometry
{
    public double H { get; init; }
    public double B { get; init; }
    public double s { get; init; }
    public double t { get; init; }
    public double A { get; init; }
    public double P { get; init; }
    public double Iz { get; init; }
    public double Iy { get; init; }
    public double Ix { get; init; }
    public double Wz { get; init; }
    public double Wy { get; init; }
    public double Wx { get; init; }
    public double Sz { get; init; }
    public double Sy { get; init; }
    public double iz { get; init; }
    public double iy { get; init; }
    public double xo { get; init; }
    public double yo { get; init; }
}
