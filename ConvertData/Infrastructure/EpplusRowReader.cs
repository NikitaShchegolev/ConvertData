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

            var header = new List<string>();
            for (int c = startCol; c <= endCol; c++)
                header.Add((ws.Cells[startRow, c].Text ?? "").Trim());

            int idxName = IndexOfHeader(header, "Name");
            int idxCode = IndexOfHeader(header, "CONNECTION_CODE");
            int idxProfile = IndexOfHeader(header, "Profile");
            int idxN = IndexOfHeader(header, "N");
            int idxQ = IndexOfHeader(header, "Q");
            int idxQo = IndexOfHeader(header, "Qo");
            int idxT = IndexOfHeader(header, "T");

            if (idxName < 0 || idxCode < 0 || idxProfile < 0 || idxN < 0 || idxQ < 0 || idxQo < 0 || idxT < 0)
                throw new InvalidDataException("Cannot find required headers in first row of worksheet");

            int colName = startCol + idxName;
            int colCode = startCol + idxCode;
            int colProfile = startCol + idxProfile;
            int colN = startCol + idxN;
            int colQ = startCol + idxQ;
            int colQo = startCol + idxQo;
            int colT = startCol + idxT;

            var list = new List<Row>();
            for (int r = startRow + 1; r <= endRow; r++)
            {
                string name = (ws.Cells[r, colName].Text ?? "").Trim();
                string code = (ws.Cells[r, colCode].Text ?? "").Trim();
                string profile = (ws.Cells[r, colProfile].Text ?? "").Trim();
                string n = (ws.Cells[r, colN].Text ?? "").Trim();
                string q = (ws.Cells[r, colQ].Text ?? "").Trim();
                string qo = (ws.Cells[r, colQo].Text ?? "").Trim();
                string t = (ws.Cells[r, colT].Text ?? "").Trim();

                if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(code))
                    continue;

                list.Add(MapBasic(name, code, profile, n, q, qo, t));
            }

            return list;
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
        /// Преобразует набор строк в доменную модель `Row`.
        /// </summary>
        private static Row MapBasic(string name, string code, string profile, string n, string q, string qo, string t)
        {
            var nInt = ParseInt(n);
            var qInt = ParseInt(q);
            var qoInt = ParseInt(qo);
            var tInt = ParseInt(t);

            return new Row
            {
                Name = name,
                CONNECTION_CODE = code,
                Profile = profile,
                Nt = nInt,
                Nc = nInt,
                N = nInt,
                Qo = qoInt,
                Q = qInt,
                T = tInt,
                M = 0,
                Mneg = 0.0,
                Mo = 0.0,
                Alpha = 0.0,
                Beta = 0.0,
                Gamma = 0.0,
                Delta = 0.0,
                Epsilon = 0.0,
                Lambda = 0.0
            };
        }

        /// <summary>
        /// Парсит целое число из строки (InvariantCulture / ru-RU), иначе возвращает 0.
        /// </summary>
        private static int ParseInt(string s)
        {
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                return v;

            if (int.TryParse(s, NumberStyles.Integer, new CultureInfo("ru-RU"), out v))
                return v;

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
    }
}
