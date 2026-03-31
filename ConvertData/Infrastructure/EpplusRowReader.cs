using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ConvertData.Application;
using ConvertData.Domain;
using ConvertData.Infrastructure.Interop;
using ConvertData.Infrastructure.Parsing;
using OfficeOpenXml;

namespace ConvertData.Infrastructure
{
    internal sealed class EpplusRowReader : IRowReader
    {
        public List<Row> Read(string path)
        {
            var format = ExcelFileSignature.Detect(path);
            if (format == ExcelFileFormat.ZipXlsx)
                return ReadXlsxWithEpplus(path);

            var tmpXlsx = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(path) + "_converted_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                ExcelXlsConverter.ConvertXlsToXlsxViaExcel(path, tmpXlsx);
                if (!File.Exists(tmpXlsx))
                    throw new InvalidDataException("Failed to convert .xls to .xlsx (temporary file not created)");

                return ReadXlsxWithEpplus(tmpXlsx);
            }
            finally
            {
                try { if (File.Exists(tmpXlsx)) File.Delete(tmpXlsx); } catch { }
            }
        }

        private static List<Row> ReadXlsxWithEpplus(string path)
        {
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets.FirstOrDefault();
            if (ws == null || ws.Dimension == null)
                return new List<Row>();

            int startRow = ws.Dimension.Start.Row;
            int endRow = ws.Dimension.End.Row;
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;

            int headerRow = FindHeaderRow(ws, startRow, Math.Min(endRow, startRow + 30), startCol, endCol);

            var header = new List<string>();
            for (int c = startCol; c <= endCol; c++)
                header.Add(HeaderUtils.NormalizeHeader((ws.Cells[headerRow, c].Text ?? "").Trim()));

            var map = ExcelHeaderResolver.Resolve(header);

            if (!map.IsMainTable && !map.IsProfileTable)
            {
                ExcelHeaderResolver.ApplyProfileFallback(map, header);

                if (!map.IsProfileTable)
                    throw new InvalidDataException("Cannot find required headers in worksheet");
            }

            var list = new List<Row>();
            int firstDataRow = headerRow + 1;

            for (int r = firstDataRow; r <= endRow; r++)
            {
                if (map.IsMainTable)
                {
                    string code = GetCell(ws, r, startCol + map.IdxCode);
                    if (string.IsNullOrWhiteSpace(code))
                        continue;

                    list.Add(RowMapper.MapMainRow(
                        GetCell(ws, r, startCol + map.IdxName),
                        code,
                        GetCell(ws, r, startCol + map.IdxProfile),
                        GetCell(ws, r, map.IdxProfileColumn >= 0 ? startCol + map.IdxProfileColumn : null),
                        GetCell(ws, r, map.IdxH >= 0 ? startCol + map.IdxH : null),
                        GetCell(ws, r, map.IdxB >= 0 ? startCol + map.IdxB : null),
                        GetCell(ws, r, map.Idxs >= 0 ? startCol + map.Idxs : null),
                        GetCell(ws, r, map.Idxt >= 0 ? startCol + map.Idxt : null),
                        GetCell(ws, r, map.IdxNt >= 0 ? startCol + map.IdxNt : null),
                        GetCell(ws, r, map.IdxQy >= 0 ? startCol + map.IdxQy : null),
                        GetCell(ws, r, map.IdxQz >= 0 ? startCol + map.IdxQz : null),
                        GetCell(ws, r, map.IdxT >= 0 ? startCol + map.IdxT : null),
                        GetCell(ws, r, map.IdxNc >= 0 ? startCol + map.IdxNc : null),
                        GetCell(ws, r, map.IdxN >= 0 ? startCol + map.IdxN : null),
                        GetCell(ws, r, map.IdxMy >= 0 ? startCol + map.IdxMy : null),
                        GetCell(ws, r, map.IdxVariable >= 0 ? startCol + map.IdxVariable : null),
                        GetCell(ws, r, map.IdxSj >= 0 ? startCol + map.IdxSj : null),
                        GetCell(ws, r, map.IdxSjo >= 0 ? startCol + map.IdxSjo : null),
                        GetCell(ws, r, map.IdxMneg >= 0 ? startCol + map.IdxMneg : null),
                        GetCell(ws, r, map.IdxMz >= 0 ? startCol + map.IdxMz : null),
                        GetCell(ws, r, map.IdxMx >= 0 ? startCol + map.IdxMx : null),
                        GetCell(ws, r, map.IdxMw >= 0 ? startCol + map.IdxMw : null),
                        GetCell(ws, r, map.IdxAlpha >= 0 ? startCol + map.IdxAlpha : null),
                        GetCell(ws, r, map.IdxBeta >= 0 ? startCol + map.IdxBeta : null),
                        GetCell(ws, r, map.IdxGamma >= 0 ? startCol + map.IdxGamma : null),
                        GetCell(ws, r, map.IdxDelta >= 0 ? startCol + map.IdxDelta : null),
                        GetCell(ws, r, map.IdxEpsilon >= 0 ? startCol + map.IdxEpsilon : null),
                        GetCell(ws, r, map.IdxLambda >= 0 ? startCol + map.IdxLambda : null)));
                }
                else
                {
                    string profile = GetCell(ws, r, startCol + map.IdxProfile);
                    if (string.IsNullOrWhiteSpace(profile))
                        continue;

                    list.Add(RowMapper.MapProfileRow(
                        profile,
                        GetCell(ws, r, map.IdxH >= 0 ? startCol + map.IdxH : null),
                        GetCell(ws, r, map.IdxB >= 0 ? startCol + map.IdxB : null),
                        GetCell(ws, r, map.Idxs >= 0 ? startCol + map.Idxs : null),
                        GetCell(ws, r, map.Idxt >= 0 ? startCol + map.Idxt : null)));
                }
            }

            if (map.IsMainTable)
                MergeAdditionalSheets(package, ws, list);

            return list;
        }

        private static string GetCell(ExcelWorksheet ws, int row, int? col)
        {
            if (col == null)
                return "";
            return (ws.Cells[row, col.Value].Text ?? "").Trim();
        }

        private static int FindHeaderRow(ExcelWorksheet ws, int fromRow, int toRow, int startCol, int endCol)
        {
            for (int r = fromRow; r <= toRow; r++)
            {
                var tokens = new List<string>();
                for (int c = startCol; c <= endCol; c++)
                    tokens.Add(HeaderUtils.NormalizeHeader((ws.Cells[r, c].Text ?? "").Trim()));

                bool hasProfile = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "ProfileBeam", "Профиль" }) >= 0;
                bool hasH = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "Beam_H", "Н" }) >= 0;
                bool hasB = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "Beam_B", "В" }) >= 0;
                bool hass = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "Beam_s", "S" }) >= 0;
                bool hast = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "Beam_t", "T" }) >= 0;

                bool hasMain = HeaderUtils.IndexOfHeaderAny(tokens, KeyColumnHeaders) >= 0
                    && HeaderUtils.IndexOfHeader(tokens, "Name") >= 0
                    && hasProfile;

                if (hasMain || (hasProfile && hasH && hasB && hass && hast))
                    return r;
            }

            return fromRow;
        }

        #region Merge additional sheets (geometry, bolts, weld)

        private static readonly string[] KeyColumnHeaders =
            ["CONNECTION_CODE", "Connection_Code", "Code", "Код"];

        private static readonly Dictionary<string, Action<Row, string>> GeometryColumnMap =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["H"] = (r, v) => { var d = NumericParser.ParseDouble(v); r.Plate_H = d; r.Flange_H = d; },
                ["B"] = (r, v) => { var d = NumericParser.ParseDouble(v); r.Plate_B = d; r.Flange_B = d; },
                ["tp"] = (r, v) => { var d = NumericParser.ParseDouble(v); r.Plate_t = d; r.Flange_t = d; },
                ["Lb"] = (r, v) => r.Flange_Lb = NumericParser.ParseDouble(v),
                ["tbp"] = (r, v) => r.Stiff_tbp = NumericParser.ParseDouble(v),
                ["tg"] = (r, v) => r.Stiff_tg = NumericParser.ParseDouble(v),
                ["tf"] = (r, v) => r.Stiff_tf = NumericParser.ParseDouble(v),
                ["Lh"] = (r, v) => r.Stiff_Lh = NumericParser.ParseDouble(v),
                ["Hh"] = (r, v) => r.Stiff_Hh = NumericParser.ParseDouble(v),
                ["tr1"] = (r, v) => r.Stiff_tr1 = NumericParser.ParseDouble(v),
                ["tr2"] = (r, v) => r.Stiff_tr2 = NumericParser.ParseDouble(v),
                ["twp"] = (r, v) => r.Stiff_twp = NumericParser.ParseDouble(v),
            };

        private static readonly Dictionary<string, Action<Row, string>> WeldColumnMap =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["kf1"] = (r, v) => r.kf1 = NumericParser.ParseInt(v),
                ["kf2"] = (r, v) => r.kf2 = NumericParser.ParseInt(v),
                ["kf3"] = (r, v) => r.kf3 = NumericParser.ParseInt(v),
                ["kf4"] = (r, v) => r.kf4 = NumericParser.ParseInt(v),
                ["kf5"] = (r, v) => r.kf5 = NumericParser.ParseInt(v),
                ["kf6"] = (r, v) => r.kf6 = NumericParser.ParseInt(v),
                ["kf7"] = (r, v) => r.kf7 = NumericParser.ParseInt(v),
                ["kf8"] = (r, v) => r.kf8 = NumericParser.ParseInt(v),
                ["kf9"] = (r, v) => r.kf9 = NumericParser.ParseInt(v),
                ["kf10"] = (r, v) => r.kf10 = NumericParser.ParseInt(v),
            };

        private static readonly Dictionary<string, Action<Row, string>> BoltsColumnMap = BuildBoltsColumnMap();

        private static Dictionary<string, Action<Row, string>> BuildBoltsColumnMap()
        {
            var map = new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
            {
                ["Option"] = (r, v) => r.OptionBolts = NumericParser.ParseInt(v),
                ["F"] = (r, v) => { r.F = NumericParser.ParseInt(v); r.N_Rows = 1; },
                ["Nb"] = (r, v) => r.Bolts_Nb = NumericParser.ParseInt(v),
                ["d1"] = (r, v) =>
                {
                    EnsureBolts(r, 1);
                    r.CoordinatesBolts[0].X = NumericParser.ParseInt(v);
                },
                ["d2"] = (r, v) =>
                {
                    EnsureBolts(r, 2);
                    r.CoordinatesBolts[1].X = NumericParser.ParseInt(v);
                    if (r.N_Rows < 2) r.N_Rows = 2;
                },
                ["e1"] = (r, v) => r.e1 = NumericParser.ParseInt(v),
                ["p1"] = (r, v) => r.p1 = NumericParser.ParseInt(v),
                ["p2"] = (r, v) => r.p2= NumericParser.ParseInt(v),
                ["p3"] = (r, v) => r.p3= NumericParser.ParseInt(v),
                ["p4"] = (r, v) => r.p4= NumericParser.ParseInt(v),
                ["p5"] = (r, v) => r.p5= NumericParser.ParseInt(v),
                ["p6"] = (r, v) => r.p6= NumericParser.ParseInt(v),
                ["p7"] = (r, v) => r.p7= NumericParser.ParseInt(v),
                ["p8"] = (r, v) => r.p8= NumericParser.ParseInt(v),
                ["p9"] = (r, v) => r.p9= NumericParser.ParseInt(v),
                ["p10"] = (r, v) => r.p10 = NumericParser.ParseInt(v),
                ["Марка опорного столика"] = (r, v) => r.TableBrand = v
            };
            return map;
        }
        /// <summary>
        /// Метод для присвоения координат болтов, гарантируя, 
        /// что список CoordinatesBolts имеет достаточное количество 
        /// элементов для доступа по индексу.
        /// </summary>
        /// <param name="r"></param>
        /// <param name="count"></param>
        private static void EnsureBolts(Row r, int count)
        {
            while (r.CoordinatesBolts.Count < count)
                r.CoordinatesBolts.Add(new CoordinatesBolts(0, 0, 0));
        }

        private static void MergeAdditionalSheets(ExcelPackage package, ExcelWorksheet mainWs, List<Row> list)
        {
            if (list.Count == 0)
                return;

            var codeLookup = new Dictionary<string, Row>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in list)
            {
                if (!string.IsNullOrWhiteSpace(row.CONNECTION_CODE) && !codeLookup.ContainsKey(row.CONNECTION_CODE))
                    codeLookup[row.CONNECTION_CODE] = row;
            }

            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws == mainWs || ws.Dimension == null)
                    continue;

                var sheetName = (ws.Name ?? "").Trim();
                if (string.Equals(sheetName, "geometry", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, GeometryColumnMap, codeLookup, list);
                else if (string.Equals(sheetName, "bolts", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, BoltsColumnMap, codeLookup, list);
                else if (string.Equals(sheetName, "weld", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, WeldColumnMap, codeLookup, list);
            }
        }

        private static void MergeSheet(
            ExcelWorksheet ws,
            Dictionary<string, Action<Row, string>> propertyMap,
            Dictionary<string, Row> codeLookup,
            List<Row> list)
        {
            int startRow = ws.Dimension.Start.Row;
            int endRow = ws.Dimension.End.Row;
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;

            var headers = new List<string>();
            for (int c = startCol; c <= endCol; c++)
                headers.Add(HeaderUtils.NormalizeHeader((ws.Cells[startRow, c].Text ?? "").Trim()));

            int keyCol = -1;
            for (int i = 0; i < headers.Count; i++)
            {
                foreach (var name in KeyColumnHeaders)
                {
                    if (string.Equals(headers[i], name, StringComparison.OrdinalIgnoreCase))
                    {
                        keyCol = i;
                        break;
                    }
                }
                if (keyCol >= 0) break;
            }

            var colMappings = new List<(int colIdx, Action<Row, string> setter)>();
            for (int i = 0; i < headers.Count; i++)
            {
                if (i == keyCol || string.IsNullOrWhiteSpace(headers[i]))
                    continue;
                if (propertyMap.TryGetValue(headers[i], out var setter))
                    colMappings.Add((i, setter));
            }

            if (colMappings.Count == 0)
                return;

            for (int r = startRow + 1; r <= endRow; r++)
            {
                Row? target = null;

                if (keyCol >= 0)
                {
                    var key = (ws.Cells[r, startCol + keyCol].Text ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(key))
                        codeLookup.TryGetValue(key, out target);
                }

                if (target == null)
                {
                    int idx = r - startRow - 1;
                    if (idx >= 0 && idx < list.Count)
                        target = list[idx];
                }

                if (target == null)
                    continue;

                foreach (var (colIdx, setter) in colMappings)
                {
                    var text = (ws.Cells[r, startCol + colIdx].Text ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                        setter(target, text);
                }
            }
        }
        #endregion
    }
}
