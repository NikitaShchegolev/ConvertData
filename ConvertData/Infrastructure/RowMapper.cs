using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

internal static class RowMapper
{
    public static Row MapMainRow(
        string name,
        string code,
        string profile,
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
            ProfileBeam = profile,
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
