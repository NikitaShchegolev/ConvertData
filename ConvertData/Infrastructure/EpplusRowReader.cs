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
                        GetCell(ws, r, map.IdxH >= 0 ? startCol + map.IdxH : null),
                        GetCell(ws, r, map.IdxB >= 0 ? startCol + map.IdxB : null),
                        GetCell(ws, r, map.Idxs >= 0 ? startCol + map.Idxs : null),
                        GetCell(ws, r, map.Idxt >= 0 ? startCol + map.Idxt : null),
                        GetCell(ws, r, map.IdxNt >= 0 ? startCol + map.IdxNt : null),
                        GetCell(ws, r, map.IdxQ >= 0 ? startCol + map.IdxQ : null),
                        GetCell(ws, r, map.IdxQo >= 0 ? startCol + map.IdxQo : null),
                        GetCell(ws, r, map.IdxT >= 0 ? startCol + map.IdxT : null),
                        GetCell(ws, r, map.IdxNc >= 0 ? startCol + map.IdxNc : null),
                        GetCell(ws, r, map.IdxN >= 0 ? startCol + map.IdxN : null),
                        GetCell(ws, r, map.IdxM >= 0 ? startCol + map.IdxM : null),
                        GetCell(ws, r, map.IdxVariable >= 0 ? startCol + map.IdxVariable : null),
                        GetCell(ws, r, map.IdxSj >= 0 ? startCol + map.IdxSj : null),
                        GetCell(ws, r, map.IdxSjo >= 0 ? startCol + map.IdxSjo : null),
                        GetCell(ws, r, map.IdxMneg >= 0 ? startCol + map.IdxMneg : null),
                        GetCell(ws, r, map.IdxMo >= 0 ? startCol + map.IdxMo : null),
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

                bool hasProfile = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "Profile", "Профиль" }) >= 0;
                bool hasH = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "H", "Н" }) >= 0;
                bool hasB = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "B", "В" }) >= 0;
                bool hass = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "s", "S" }) >= 0;
                bool hast = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "t", "T" }) >= 0;

                bool hasMain = HeaderUtils.IndexOfHeaderAny(tokens, new[] { "CONNECTION_CODE", "Connection_Code", "Code", "Код" }) >= 0
                    && HeaderUtils.IndexOfHeader(tokens, "Name") >= 0
                    && hasProfile;

                if (hasMain || (hasProfile && hasH && hasB && hass && hast))
                    return r;
            }

            return fromRow;
        }
    }
}
