namespace ConvertData.Domain;

internal sealed class Row
{
    public string Name { get; set; } = "";
    public string CONNECTION_CODE { get; set; } = "";
    public string Profile { get; set; } = "";

    public double H { get; set; }
    public double B { get; set; }
    public double s { get; set; }
    public double t { get; set; }

    public int Nt { get; set; }
    public int Nc { get; set; }
    public int N { get; set; }
    public int Qo { get; set; }
    public int Q { get; set; }
    public int T { get; set; }
    public int M { get; set; }

    public int variable { get; set; }
    public int Sj { get; set; }
    public int Sjo { get; set; }

    public double Mneg { get; set; }
    public double Mo { get; set; }

    public double Alpha { get; set; }
    public double Beta { get; set; }
    public double Gamma { get; set; }
    public double Delta { get; set; }
    public double Epsilon { get; set; }
    public double Lambda { get; set; }
}
