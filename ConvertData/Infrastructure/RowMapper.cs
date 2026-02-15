using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

internal static class RowMapper
{
    public static Row MapMainRow(
        string name,
        string code,
        string profile,
        string h,
        string b,
        string s,
        string tGeom,
        string nt,
        string q,
        string qo,
        string t,
        string nc,
        string n,
        string m,
        string variable,
        string sj,
        string sjo,
        string mneg,
        string mo,
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
            Profile = profile,
            H = NumericParser.ParseDouble(h),
            B = NumericParser.ParseDouble(b),
            s = NumericParser.ParseDouble(s),
            t = NumericParser.ParseDouble(tGeom),
            Nt = NumericParser.ParseInt(nt),
            Nc = NumericParser.ParseInt(nc),
            N = NumericParser.ParseInt(n),
            Qo = NumericParser.ParseInt(qo),
            Q = NumericParser.ParseInt(q),
            T = NumericParser.ParseInt(t),
            M = NumericParser.ParseInt(m),
            variable = NumericParser.ParseInt(variable),
            Sj = NumericParser.ParseInt(sj),
            Sjo = NumericParser.ParseInt(sjo),
            Mneg = NumericParser.ParseDouble(mneg),
            Mo = NumericParser.ParseDouble(mo),
            Alpha = NumericParser.ParseDouble(alpha),
            Beta = NumericParser.ParseDouble(beta),
            Gamma = NumericParser.ParseDouble(gamma),
            Delta = NumericParser.ParseDouble(delta),
            Epsilon = NumericParser.ParseDouble(epsilon),
            Lambda = NumericParser.ParseDouble(lambda)
        };
    }

    public static Row MapProfileRow(string profile, string h, string b, string s, string t)
    {
        return new Row
        {
            Profile = profile,
            H = NumericParser.ParseDouble(h),
            B = NumericParser.ParseDouble(b),
            s = NumericParser.ParseDouble(s),
            t = NumericParser.ParseDouble(t)
        };
    }
}
