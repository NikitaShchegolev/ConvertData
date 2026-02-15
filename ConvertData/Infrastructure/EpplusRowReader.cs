using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ConvertData.Application;
using ConvertData.Domain;
using OfficeOpenXml;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Reader для Excel-файлов.
    ///
    /// Поддержка:
    /// - `.xlsx` (zip) читается напрямую через EPPlus;
    /// - бинарный `.xls` (OLE) при необходимости конвертируется в временный `.xlsx` через установленный Microsoft Excel (COM).
    /// </summary>
    internal sealed class EpplusRowReader : IRowReader
    {
        /// <summary>
        /// Считывает Excel-файл и возвращает строки в виде доменных объектов.
        /// </summary>
        /// <param name="path">Путь к Excel файлу.</param>
        public List<Row> Read(string path)
        {
            return ReadXlsxOrXlsViaExcelInterop(path);
        }

        /// <summary>
        /// Определяет формат Excel по сигнатуре файла.
        /// Если это `.xlsx` — читает напрямую.
        /// Если это бинарный `.xls` — конвертирует во временный `.xlsx` через Excel и читает.
        /// </summary>
        private static List<Row> ReadXlsxOrXlsViaExcelInterop(string path)
        {
            byte[] header = new byte[8];
            using (var fsSig = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                int read = fsSig.Read(header, 0, header.Length);
                if (read < 2)
                    throw new InvalidDataException("File is too small");
            }

            bool isZip = header[0] == (byte)'P' && header[1] == (byte)'K';
            if (isZip)
                return ReadXlsxWithEpplus(path);

            bool isOle = header.Length >= 8
                && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0
                && header[4] == 0xA1 && header[5] == 0xB1 && header[6] == 0x1A && header[7] == 0xE1;

            if (!isOle)
                throw new InvalidDataException("Unknown Excel format (not zip/xlsx and not OLE/xls)");

            var tmpXlsx = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(path) + "_converted_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                ConvertXlsToXlsxViaExcel(path, tmpXlsx);
                if (!File.Exists(tmpXlsx))
                    throw new InvalidDataException("Failed to convert .xls to .xlsx (temporary file not created)");

                return ReadXlsxWithEpplus(tmpXlsx);
            }
            finally
            {
                try { if (File.Exists(tmpXlsx)) File.Delete(tmpXlsx); } catch { }
            }
        }

        /// <summary>
        /// Читает `.xlsx` через EPPlus: ищет первую страницу, определяет колонки по заголовкам,
        /// затем мапит строки на `Row`.
        /// </summary>
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

            // Ищем строку заголовка (иногда первая строка может быть пустой/с названием таблицы)
            int headerRow = FindHeaderRow(ws, startRow, Math.Min(endRow, startRow + 30), startCol, endCol);

            var header = new List<string>();
            for (int c = startCol; c <= endCol; c++)
                header.Add(NormalizeHeader((ws.Cells[headerRow, c].Text ?? "").Trim()));

            int idxH = IndexOfHeaderAny(header, new[] { "H", "Н" });
            int idxB = IndexOfHeaderAny(header, new[] { "B", "В" });
            int idxs = IndexOfHeaderAny(header, new[] { "s", "S" });
            int idxt = IndexOfHeaderAny(header, new[] { "t", "T" });

            int idxName = IndexOfHeader(header, "Name");
            int idxCode = IndexOfHeaderAny(header, new[] { "CONNECTION_CODE", "Connection_Code", "Code", "Код" });
            int idxProfile = IndexOfHeaderAny(header, new[] { "Profile", "Профиль" });

            int idxNt = IndexOfHeader(header, "Nt");
            int idxQ = IndexOfHeader(header, "Q");
            int idxQo = IndexOfHeader(header, "Qo");
            int idxT = IndexOfHeader(header, "T");
            int idxNc = IndexOfHeader(header, "Nc");
            int idxN = IndexOfHeader(header, "N");
            int idxM = IndexOfHeader(header, "M");
            int idxVariable = IndexOfHeaderAny(header, new[] { "variable", "Variable" });
            int idxSj = IndexOfHeader(header, "Sj");
            int idxSjo = IndexOfHeader(header, "Sjo");
            int idxMneg = IndexOfHeader(header, "Mneg");
            int idxMo = IndexOfHeader(header, "Mo");

            int idxAlpha = IndexOfHeader(header, "α");
            if (idxAlpha < 0) idxAlpha = IndexOfHeader(header, "Alpha");
            int idxBeta = IndexOfHeader(header, "β");
            if (idxBeta < 0) idxBeta = IndexOfHeader(header, "Beta");
            int idxGamma = IndexOfHeader(header, "γ");
            if (idxGamma < 0) idxGamma = IndexOfHeader(header, "Gamma");
            int idxDelta = IndexOfHeader(header, "δ");
            if (idxDelta < 0) idxDelta = IndexOfHeader(header, "Delta");
            int idxEpsilon = IndexOfHeader(header, "ε");
            if (idxEpsilon < 0) idxEpsilon = IndexOfHeader(header, "Epsilon");
            int idxLambda = IndexOfHeader(header, "λ");
            if (idxLambda < 0) idxLambda = IndexOfHeader(header, "Lambda");

            // Fallback: иногда греческие буквы теряются и приходят как "?".
            // В этом случае подхватываем 6 колонок после Mo.
            if (idxMo >= 0 && (idxAlpha < 0 || idxBeta < 0 || idxGamma < 0 || idxDelta < 0 || idxEpsilon < 0 || idxLambda < 0))
            {
                var qMarks = header
                    .Select((h, i) => new { h, i })
                    .Where(x => x.h == "?")
                    .Select(x => x.i)
                    .ToList();

                int baseIdx = idxMo + 1;
                if (baseIdx < header.Count && header.Count - baseIdx >= 6)
                {
                    if (idxAlpha < 0) idxAlpha = baseIdx + 0;
                    if (idxBeta < 0) idxBeta = baseIdx + 1;
                    if (idxGamma < 0) idxGamma = baseIdx + 2;
                    if (idxDelta < 0) idxDelta = baseIdx + 3;
                    if (idxEpsilon < 0) idxEpsilon = baseIdx + 4;
                    if (idxLambda < 0) idxLambda = baseIdx + 5;
                }
                else if (qMarks.Count >= 6)
                {
                    if (idxAlpha < 0) idxAlpha = qMarks[0];
                    if (idxBeta < 0) idxBeta = qMarks[1];
                    if (idxGamma < 0) idxGamma = qMarks[2];
                    if (idxDelta < 0) idxDelta = qMarks[3];
                    if (idxEpsilon < 0) idxEpsilon = qMarks[4];
                    if (idxLambda < 0) idxLambda = qMarks[5];
                }
            }

            // Поддержка двух форматов:
            // 1) Основные таблицы: обязательны Name/CONNECTION_CODE/Profile
            // 2) Справочник профилей: достаточно Profile + H/B/s/t
            bool isMainTable = idxName >= 0 && idxCode >= 0 && idxProfile >= 0;
            bool isProfileTable = idxProfile >= 0 && idxH >= 0 && idxB >= 0 && idxs >= 0 && idxt >= 0;

            if (!isMainTable && !isProfileTable)
            {
                // Fallback для справочника: заголовки могут отсутствовать или быть нераспознаны.
                // 1) Если есть колонка Profile, считаем что следующие 4 колонки — H,B,s,t.
                if (idxProfile >= 0)
                {
                    if (idxH < 0) idxH = idxProfile + 1;
                    if (idxB < 0) idxB = idxProfile + 2;
                    if (idxs < 0) idxs = idxProfile + 3;
                    if (idxt < 0) idxt = idxProfile + 4;
                    isProfileTable = idxProfile >= 0 && idxH < header.Count && idxB < header.Count && idxs < header.Count && idxt < header.Count;
                }
                else
                {
                    // 2) Если даже Profile не нашли — берём первые 5 колонок.
                    idxProfile = 0;
                    idxH = 1;
                    idxB = 2;
                    idxs = 3;
                    idxt = 4;
                    isProfileTable = header.Count >= 5;
                }

                if (!isProfileTable)
                    throw new InvalidDataException("Cannot find required headers in worksheet");
            }

            int? colH = idxH >= 0 ? startCol + idxH : null;
            int? colB = idxB >= 0 ? startCol + idxB : null;
            int? cols = idxs >= 0 ? startCol + idxs : null;
            int? colt = idxt >= 0 ? startCol + idxt : null;

            int? colName = idxName >= 0 ? startCol + idxName : null;
            int? colCode = idxCode >= 0 ? startCol + idxCode : null;
            int colProfile = startCol + idxProfile;

            int? colNt = idxNt >= 0 ? startCol + idxNt : null;
            int? colQ = idxQ >= 0 ? startCol + idxQ : null;
            int? colQo = idxQo >= 0 ? startCol + idxQo : null;
            int? colT = idxT >= 0 ? startCol + idxT : null;
            int? colNc = idxNc >= 0 ? startCol + idxNc : null;
            int? colN = idxN >= 0 ? startCol + idxN : null;
            int? colM = idxM >= 0 ? startCol + idxM : null;
            int? colVariable = idxVariable >= 0 ? startCol + idxVariable : null;
            int? colSj = idxSj >= 0 ? startCol + idxSj : null;
            int? colSjo = idxSjo >= 0 ? startCol + idxSjo : null;
            int? colMneg = idxMneg >= 0 ? startCol + idxMneg : null;
            int? colMo = idxMo >= 0 ? startCol + idxMo : null;
            int? colAlpha = idxAlpha >= 0 ? startCol + idxAlpha : null;
            int? colBeta = idxBeta >= 0 ? startCol + idxBeta : null;
            int? colGamma = idxGamma >= 0 ? startCol + idxGamma : null;
            int? colDelta = idxDelta >= 0 ? startCol + idxDelta : null;
            int? colEpsilon = idxEpsilon >= 0 ? startCol + idxEpsilon : null;
            int? colLambda = idxLambda >= 0 ? startCol + idxLambda : null;

            var list = new List<Row>();
            int firstDataRow = headerRow + 1;

            for (int r = firstDataRow; r <= endRow; r++)
            {
                if (isMainTable)
                {
                    string code = (ws.Cells[r, colCode!.Value].Text ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(code))
                        continue;

                    string name = (ws.Cells[r, colName!.Value].Text ?? "").Trim();
                    string profile = (ws.Cells[r, colProfile].Text ?? "").Trim();

                    string hStr = GetCell(ws, r, colH);
                    string bStr = GetCell(ws, r, colB);
                    string sStr = GetCell(ws, r, cols);
                    string tgeomStr = GetCell(ws, r, colt);

                    string ntStr = GetCell(ws, r, colNt);
                    string qStr = GetCell(ws, r, colQ);
                    string qoStr = GetCell(ws, r, colQo);
                    string tStr = GetCell(ws, r, colT);
                    string ncStr = GetCell(ws, r, colNc);
                    string nStr = GetCell(ws, r, colN);
                    string mStr = GetCell(ws, r, colM);
                    string variableStr = GetCell(ws, r, colVariable);
                    string sjStr = GetCell(ws, r, colSj);
                    string sjoStr = GetCell(ws, r, colSjo);
                    string mnegStr = GetCell(ws, r, colMneg);
                    string moStr = GetCell(ws, r, colMo);
                    string alphaStr = GetCell(ws, r, colAlpha);
                    string betaStr = GetCell(ws, r, colBeta);
                    string gammaStr = GetCell(ws, r, colGamma);
                    string deltaStr = GetCell(ws, r, colDelta);
                    string epsilonStr = GetCell(ws, r, colEpsilon);
                    string lambdaStr = GetCell(ws, r, colLambda);

                    list.Add(Map19(name, code, profile, hStr, bStr, sStr, tgeomStr, ntStr, qStr, qoStr, tStr, ncStr, nStr, mStr, variableStr, sjStr, sjoStr, mnegStr, moStr, alphaStr, betaStr, gammaStr, deltaStr, epsilonStr, lambdaStr));
                }
                else
                {
                    // Справочник Profile.xls: Profile/H/B/s/t
                    string profile = (ws.Cells[r, colProfile].Text ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(profile))
                        continue;

                    string hStr = GetCell(ws, r, colH);
                    string bStr = GetCell(ws, r, colB);
                    string sStr = GetCell(ws, r, cols);
                    string tStr = GetCell(ws, r, colt);

                    list.Add(new Row
                    {
                        Profile = profile,
                        H = ParseDouble(hStr),
                        B = ParseDouble(bStr),
                        s = ParseDouble(sStr),
                        t = ParseDouble(tStr)
                    });
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

        private static Row Map19(
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
            var hDouble = ParseDouble(h);
            var bDouble = ParseDouble(b);
            var sDouble = ParseDouble(s);
            var tGeomDouble = ParseDouble(tGeom);

            var ntInt = ParseInt(nt);
            var qInt = ParseInt(q);
            var qoInt = ParseInt(qo);
            var tInt = ParseInt(t);
            var ncInt = ParseInt(nc);
            var nInt = ParseInt(n);
            var mInt = ParseInt(m);
            var variableInt = ParseInt(variable);
            var sjInt = ParseInt(sj);
            var sjoInt = ParseInt(sjo);

            var mnegDouble = ParseDouble(mneg);
            var moDouble = ParseDouble(mo);
            var alphaDouble = ParseDouble(alpha);
            var betaDouble = ParseDouble(beta);
            var gammaDouble = ParseDouble(gamma);
            var deltaDouble = ParseDouble(delta);
            var epsilonDouble = ParseDouble(epsilon);
            var lambdaDouble = ParseDouble(lambda);

            return new Row
            {
                Name = name,
                CONNECTION_CODE = code,
                Profile = profile,
                H = hDouble,
                B = bDouble,
                s = sDouble,
                t = tGeomDouble,
                Nt = ntInt,
                Nc = ncInt,
                N = nInt,
                Qo = qoInt,
                Q = qInt,
                T = tInt,
                M = mInt,
                variable = variableInt,
                Sj = sjInt,
                Sjo = sjoInt,
                Mneg = mnegDouble,
                Mo = moDouble,
                Alpha = alphaDouble,
                Beta = betaDouble,
                Gamma = gammaDouble,
                Delta = deltaDouble,
                Epsilon = epsilonDouble,
                Lambda = lambdaDouble
            };
        }

        private static readonly CultureInfo RuCulture = new("ru-RU");

        private static double ParseDouble(string s)
        {
            if (s.Contains(','))
            {
                if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, RuCulture, out var vr))
                    return vr;
            }

            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                return v;

            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, RuCulture, out v))
                return v;

            return 0.0;
        }

        /// <summary>
        /// Конвертирует бинарный `.xls` в `.xlsx` через установленный Microsoft Excel (COM automation).
        /// Нужен установленный Excel и Windows.
        /// </summary>
        private static void ConvertXlsToXlsxViaExcel(string xlsPath, string xlsxPath)
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new PlatformNotSupportedException("Conversion from .xls requires Microsoft Excel installed (Excel.Application COM is not available).");

            object? excel = null;
            object? workbooks = null;
            object? workbook = null;

            try
            {
                excel = Activator.CreateInstance(excelType);
                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });
                excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, excel, new object[] { false });

                workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excel, Array.Empty<object>());
                var workbooksType = workbooks!.GetType();

                workbook = workbooksType.InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    workbooks,
                    new object[] { xlsPath }
                );

                workbook!.GetType().InvokeMember(
                    "SaveAs",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    workbook,
                    new object[] { xlsxPath, 51 }
                );

                workbook.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                workbook = null;

                excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, Array.Empty<object>());
            }
            finally
            {
                try
                {
                    if (workbook != null)
                        workbook.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                }
                catch { }

                try
                {
                    if (excel != null)
                        excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, excel, Array.Empty<object>());
                }
                catch { }

                if (workbook != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                if (workbooks != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbooks);
                if (excel != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel);
            }
        }

        /// <summary>
        /// Парсит целое число из строки (InvariantCulture / ru-RU), иначе возвращает 0.
        /// </summary>
        private static int ParseInt(string s)
        {
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                return v;

            if (int.TryParse(s, NumberStyles.Integer, RuCulture, out v))
                return v;

            var d = ParseDouble(s);
            if (d != 0.0)
                return (int)Math.Round(d);

            return 0;
        }

        /// <summary>
        /// Находит индекс колонки в строке заголовков по имени (без учёта регистра).
        /// </summary>
        private static int IndexOfHeader(List<string> header, string name)
        {
            for (int i = 0; i < header.Count; i++)
            {
                if (string.Equals(header[i], name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private static int IndexOfHeaderAny(List<string> header, IEnumerable<string> names)
        {
            foreach (var n in names)
            {
                int idx = IndexOfHeader(header, n);
                if (idx >= 0)
                    return idx;
            }
            return -1;
        }

        private static int FindHeaderRow(ExcelWorksheet ws, int fromRow, int toRow, int startCol, int endCol)
        {
            for (int r = fromRow; r <= toRow; r++)
            {
                var tokens = new List<string>();
                for (int c = startCol; c <= endCol; c++)
                    tokens.Add(NormalizeHeader((ws.Cells[r, c].Text ?? "").Trim()));

                bool hasProfile = IndexOfHeaderAny(tokens, new[] { "Profile", "Профиль" }) >= 0;
                bool hasH = IndexOfHeaderAny(tokens, new[] { "H", "Н" }) >= 0;
                bool hasB = IndexOfHeaderAny(tokens, new[] { "B", "В" }) >= 0;
                bool hass = IndexOfHeaderAny(tokens, new[] { "s", "S" }) >= 0;
                bool hast = IndexOfHeaderAny(tokens, new[] { "t", "T" }) >= 0;

                bool hasMain = IndexOfHeaderAny(tokens, new[] { "CONNECTION_CODE", "Connection_Code", "Code", "Код" }) >= 0
                    && IndexOfHeader(tokens, "Name") >= 0
                    && hasProfile;

                bool hasProfileTable = hasProfile && hasH && hasB && hass && hast;

                if (hasMain || hasProfileTable)
                    return r;
            }

            return fromRow;
        }

        private static string NormalizeHeader(string h)
        {
            if (string.IsNullOrWhiteSpace(h))
                return string.Empty;

            return h.Trim();
        }
    }
}
