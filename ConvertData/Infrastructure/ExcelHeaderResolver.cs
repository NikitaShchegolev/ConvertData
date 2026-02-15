using System.Collections.Generic;
using System.Linq;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

internal sealed class ExcelColumnMap
{
    public int IdxH { get; set; } = -1;
    public int IdxB { get; set; } = -1;
    public int Idxs { get; set; } = -1;
    public int Idxt { get; set; } = -1;
    public int IdxName { get; set; } = -1;
    public int IdxCode { get; set; } = -1;
    public int IdxProfile { get; set; } = -1;
    public int IdxNt { get; set; } = -1;
    public int IdxQ { get; set; } = -1;
    public int IdxQo { get; set; } = -1;
    public int IdxT { get; set; } = -1;
    public int IdxNc { get; set; } = -1;
    public int IdxN { get; set; } = -1;
    public int IdxM { get; set; } = -1;
    public int IdxVariable { get; set; } = -1;
    public int IdxSj { get; set; } = -1;
    public int IdxSjo { get; set; } = -1;
    public int IdxMneg { get; set; } = -1;
    public int IdxMo { get; set; } = -1;
    public int IdxAlpha { get; set; } = -1;
    public int IdxBeta { get; set; } = -1;
    public int IdxGamma { get; set; } = -1;
    public int IdxDelta { get; set; } = -1;
    public int IdxEpsilon { get; set; } = -1;
    public int IdxLambda { get; set; } = -1;

    public bool IsMainTable => IdxName >= 0 && IdxCode >= 0 && IdxProfile >= 0;
    public bool IsProfileTable => IdxProfile >= 0 && IdxH >= 0 && IdxB >= 0 && Idxs >= 0 && Idxt >= 0;
}

internal static class ExcelHeaderResolver
{
    public static ExcelColumnMap Resolve(List<string> header)
    {
        var map = new ExcelColumnMap
        {
            IdxH = HeaderUtils.IndexOfHeaderAny(header, new[] { "H", "Н" }),
            IdxB = HeaderUtils.IndexOfHeaderAny(header, new[] { "B", "В" }),
            Idxs = HeaderUtils.IndexOfHeaderAny(header, new[] { "s", "S" }),
            Idxt = HeaderUtils.IndexOfHeaderAny(header, new[] { "t", "T" }),
            IdxName = HeaderUtils.IndexOfHeader(header, "Name"),
            IdxCode = HeaderUtils.IndexOfHeaderAny(header, new[] { "CONNECTION_CODE", "Connection_Code", "Code", "Код" }),
            IdxProfile = HeaderUtils.IndexOfHeaderAny(header, new[] { "Profile", "Профиль" }),
            IdxNt = HeaderUtils.IndexOfHeader(header, "Nt"),
            IdxQ = HeaderUtils.IndexOfHeader(header, "Q"),
            IdxQo = HeaderUtils.IndexOfHeader(header, "Qo"),
            IdxT = HeaderUtils.IndexOfHeader(header, "T"),
            IdxNc = HeaderUtils.IndexOfHeader(header, "Nc"),
            IdxN = HeaderUtils.IndexOfHeader(header, "N"),
            IdxM = HeaderUtils.IndexOfHeader(header, "M"),
            IdxVariable = HeaderUtils.IndexOfHeaderAny(header, new[] { "variable", "Variable" }),
            IdxSj = HeaderUtils.IndexOfHeader(header, "Sj"),
            IdxSjo = HeaderUtils.IndexOfHeader(header, "Sjo"),
            IdxMneg = HeaderUtils.IndexOfHeader(header, "Mneg"),
            IdxMo = HeaderUtils.IndexOfHeader(header, "Mo")
        };

        map.IdxAlpha = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxAlpha < 0) map.IdxAlpha = HeaderUtils.IndexOfHeader(header, "Alpha");
        map.IdxBeta = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxBeta < 0) map.IdxBeta = HeaderUtils.IndexOfHeader(header, "Beta");
        map.IdxGamma = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxGamma < 0) map.IdxGamma = HeaderUtils.IndexOfHeader(header, "Gamma");
        map.IdxDelta = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxDelta < 0) map.IdxDelta = HeaderUtils.IndexOfHeader(header, "Delta");
        map.IdxEpsilon = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxEpsilon < 0) map.IdxEpsilon = HeaderUtils.IndexOfHeader(header, "Epsilon");
        map.IdxLambda = HeaderUtils.IndexOfHeader(header, "?");
        if (map.IdxLambda < 0) map.IdxLambda = HeaderUtils.IndexOfHeader(header, "Lambda");

        ResolveGreekFallback(header, map);

        return map;
    }

    private static void ResolveGreekFallback(List<string> header, ExcelColumnMap map)
    {
        if (map.IdxMo < 0)
            return;
        if (map.IdxAlpha >= 0 && map.IdxBeta >= 0 && map.IdxGamma >= 0
            && map.IdxDelta >= 0 && map.IdxEpsilon >= 0 && map.IdxLambda >= 0)
            return;

        var qMarks = header
            .Select((h, i) => new { h, i })
            .Where(x => x.h == "?")
            .Select(x => x.i)
            .ToList();

        int baseIdx = map.IdxMo + 1;
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
