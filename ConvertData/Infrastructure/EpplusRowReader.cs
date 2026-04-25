using ConvertData.Application;
using ConvertData.Domain;
using ConvertData.Infrastructure.Interop;
using ConvertData.Infrastructure.Parsing;

using OfficeOpenXml;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Читает строки из Excel-файлов (.xls/.xlsx) используя библиотеку EPPlus.
    /// Поддерживает автоматическую конвертацию .xls в .xlsx через COM Interop.
    /// </summary>
    internal sealed class EpplusRowReader : IRowReader
    {
        /// <summary>
        /// Читает данные из Excel-файла и возвращает список объектов Row.
        /// </summary>
        /// <param name="path">Путь к Excel-файлу (.xls или .xlsx).</param>
        /// <returns>Список прочитанных строк.</returns>
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

        /// <summary>
        /// Читает XLSX-файл используя EPPlus и возвращает список Row.
        /// Обрабатывает основной лист и объединяет данные из дополнительных листов (geometry, bolts, weld).
        /// </summary>
        /// <param name="path">Путь к XLSX-файлу.</param>
        /// <returns>Список объектов Row.</returns>
        private static List<Row> ReadXlsxWithEpplus(string path)
        {
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets
                .FirstOrDefault(x => string.Equals((x.Name ?? "").Trim(), "data", StringComparison.OrdinalIgnoreCase))
                ?? package.Workbook.Worksheets.FirstOrDefault();
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
                        GetCell(ws, r, map.IdxTypeNode >= 0 ? startCol + map.IdxTypeNode : null),
                        GetCell(ws, r, map.IdxGost >= 0 ? startCol + map.IdxGost : null),
                        GetCell(ws, r, map.IdxGostColumnAndBeams >= 0 ? startCol + map.IdxGostColumnAndBeams : null),
                        GetCell(ws, r, map.IdxGostHoles >= 0 ? startCol + map.IdxGostHoles : null),
                        GetCell(ws, r, map.IdxGostBolts >= 0 ? startCol + map.IdxGostBolts : null),
                        GetCell(ws, r, map.IdxGostAnchore >= 0 ? startCol + map.IdxGostAnchore : null),
                        GetCell(ws, r, map.IdxGostWeld >= 0 ? startCol + map.IdxGostWeld : null),
                        GetCell(ws, r, map.IdxGostProfile >= 0 ? startCol + map.IdxGostProfile : null),
                        GetCell(ws, r, map.IdxTableBrand >= 0 ? startCol + map.IdxTableBrand : null),
                        GetCell(ws, r, startCol + map.IdxProfileBeam),
                        GetCell(ws, r, map.IdxProfileColumn >= 0 ? startCol + map.IdxProfileColumn : null),
                        GetCell(ws, r, map.IdxExplanations >= 0 ? startCol + map.IdxExplanations : null),
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
                        GetCell(ws, r, map.IdxMy_compression >= 0 ? startCol + map.IdxMy_compression : null),
                        GetCell(ws, r, map.IdxMy_tension >= 0 ? startCol + map.IdxMy_tension : null),
                        GetCell(ws, r, map.IdxVariable >= 0 ? startCol + map.IdxVariable : null),
                        GetCell(ws, r, map.IdxSj >= 0 ? startCol + map.IdxSj : null),
                        GetCell(ws, r, map.IdxSjo >= 0 ? startCol + map.IdxSjo : null),
                        GetCell(ws, r, map.IdxMneg >= 0 ? startCol + map.IdxMneg : null),
                        GetCell(ws, r, map.IdxMz >= 0 ? startCol + map.IdxMz : null),
                        GetCell(ws, r, map.IdxMz_compression >= 0 ? startCol + map.IdxMz_compression : null),
                        GetCell(ws, r, map.IdxMz_tension >= 0 ? startCol + map.IdxMz_tension : null),
                        GetCell(ws, r, map.IdxMx >= 0 ? startCol + map.IdxMx : null),
                        GetCell(ws, r, map.IdxMw >= 0 ? startCol + map.IdxMw : null),
                        GetCell(ws, r, map.IdxAlpha >= 0 ? startCol + map.IdxAlpha : null),
                        GetCell(ws, r, map.IdxBeta >= 0 ? startCol + map.IdxBeta : null),
                        GetCell(ws, r, map.IdxGamma >= 0 ? startCol + map.IdxGamma : null),
                        GetCell(ws, r, map.IdxDelta >= 0 ? startCol + map.IdxDelta : null),
                        GetCell(ws, r, map.IdxEpsilon >= 0 ? startCol + map.IdxEpsilon : null),
                        GetCell(ws, r, map.IdxLambda >= 0 ? startCol + map.IdxLambda : null),
                        GetCell(ws, r, map.IdxB_plate >= 0 ? startCol + map.IdxB_plate : null),
                        GetCell(ws, r, map.IdxH_plate >= 0 ? startCol + map.IdxH_plate : null),
                        GetCell(ws, r, map.IdxLws_plate >= 0 ? startCol + map.IdxLws_plate : null),
                        GetCell(ws, r, map.IdxTp_plate >= 0 ? startCol + map.IdxTp_plate : null),
                        GetCell(ws, r, map.IdxTr1_plate >= 0 ? startCol + map.IdxTr1_plate : null),
                        GetCell(ws, r, map.IdxTr2_plate >= 0 ? startCol + map.IdxTr2_plate : null),
                        GetCell(ws, r, map.IdxB_stiff >= 0 ? startCol + map.IdxB_stiff : null),
                        GetCell(ws, r, map.IdxH_stiff >= 0 ? startCol + map.IdxH_stiff : null),
                        GetCell(ws, r, map.IdxLws_stiff >= 0 ? startCol + map.IdxLws_stiff : null),
                        GetCell(ws, r, map.Idxtp_stiff >= 0 ? startCol + map.Idxtp_stiff : null),
                        GetCell(ws, r, map.Idxtr1_stiff >= 0 ? startCol + map.Idxtr1_stiff : null),
                        GetCell(ws, r, map.Idxtr2_stiff >= 0 ? startCol + map.Idxtr2_stiff : null),
                        GetCell(ws, r, map.IdF_base >= 0 ? startCol + map.IdF_base : null),
                        GetCell(ws, r, map.IdLws_base >= 0 ? startCol + map.IdLws_base : null),
                        GetCell(ws, r, map.IdLp_base >= 0 ? startCol + map.IdLp_base : null),
                        GetCell(ws, r, map.IdLs_base >= 0 ? startCol + map.IdLs_base : null),
                        GetCell(ws, r, map.IdTws_base >= 0 ? startCol + map.IdTws_base : null),
                        GetCell(ws, r, map.IdD_ws_base >= 0 ? startCol + map.IdD_ws_base : null),
                        GetCell(ws, r, map.IdD_p_base >= 0 ? startCol + map.IdD_p_base : null),
                        GetCell(ws, r, map.IdXh_base >= 0 ? startCol + map.IdXh_base : null),
                        GetCell(ws, r, map.IdxH_base >= 0 ? startCol + map.IdxH_base : null),
                        GetCell(ws, r, map.IdxB_base >= 0 ? startCol + map.IdxB_base : null),
                        GetCell(ws, r, map.IdxS_base >= 0 ? startCol + map.IdxS_base : null),
                        GetCell(ws, r, map.IdxT_base >= 0 ? startCol + map.IdxT_base : null),
                        GetCell(ws, r, map.IdNh_base_var1 >= 0 ? startCol + map.IdNh_base_var1 : null),
                        GetCell(ws, r, map.IdNh_base_var2 >= 0 ? startCol + map.IdNh_base_var2 : null),
                        GetCell(ws, r, map.IdAnchor_var_1 >= 0 ? startCol + map.IdAnchor_var_1 : null),
                        GetCell(ws, r, map.IdAnchor_var_2 >= 0 ? startCol + map.IdAnchor_var_2 : null),
                        GetCell(ws, r, map.IdAnchor_var_3 >= 0 ? startCol + map.IdAnchor_var_3 : null),
                        GetCell(ws, r, map.IdAnchor_var_4 >= 0 ? startCol + map.IdAnchor_var_4 : null),
                        GetCell(ws, r, map.IdxLp_shearKey >= 0 ? startCol + map.IdxLp_shearKey : null),
                        GetCell(ws, r, map.IdxLs_shearKey >= 0 ? startCol + map.IdxLs_shearKey : null)
                        ));
                }
                else
                {
                    string profile = GetCell(ws, r, startCol + map.IdxProfileBeam);
                    if (string.IsNullOrWhiteSpace(profile))
                        continue;

                    list.Add(RowMapper.MapProfileRow(
                        profile,
                        GetCell(ws, r, map.IdxGostProfile >= 0 ? startCol + map.IdxGostProfile : null),
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

        /// <summary>
        /// Извлекает текст из ячейки Excel.
        /// </summary>
        /// <param name="ws">Лист Excel.</param>
        /// <param name="row">Номер строки.</param>
        /// <param name="col">Номер колонки (null, если колонка не найдена).</param>
        /// <returns>Текст ячейки или пустая строка.</returns>
        private static string GetCell(ExcelWorksheet ws, int row, int? col)
        {
            if (col == null)
                return "";
            return (ws.Cells[row, col.Value].Text ?? "").Trim();
        }

        /// <summary>
        /// Ищет строку с заголовками в указанном диапазоне листа Excel.
        /// </summary>
        /// <param name="ws">Лист Excel.</param>
        /// <param name="fromRow">Начальная строка поиска.</param>
        /// <param name="toRow">Конечная строка поиска.</param>
        /// <param name="startCol">Начальная колонка.</param>
        /// <param name="endCol">Конечная колонка.</param>
        /// <returns>Номер строки с заголовками.</returns>
        private static int FindHeaderRow(ExcelWorksheet ws, int fromRow, int toRow, int startCol, int endCol)
        {
            for (int r = fromRow; r <= toRow; r++)
            {
                var tokens = new List<string>();
                for (int c = startCol; c <= endCol; c++)
                    tokens.Add(HeaderUtils.NormalizeHeader((ws.Cells[r, c].Text ?? "").Trim()));

                bool hasProfile = HeaderUtils.IndexOfHeaderAny(tokens, ["ProfileBeam", "Профиль"]) >= 0;
                bool hasH = HeaderUtils.IndexOfHeaderAny(tokens, ["Beam_H", "Н"]) >= 0;
                bool hasB = HeaderUtils.IndexOfHeaderAny(tokens, ["Beam_B", "В"]) >= 0;
                bool hass = HeaderUtils.IndexOfHeaderAny(tokens, ["Beam_s", "S"]) >= 0;
                bool hast = HeaderUtils.IndexOfHeaderAny(tokens, ["Beam_t", "T"]) >= 0;

                bool hasMain = HeaderUtils.IndexOfHeaderAny(tokens, KeyColumnHeaders) >= 0
                    && HeaderUtils.IndexOfHeader(tokens, "Name") >= 0
                    && hasProfile;

                if (hasMain || (hasProfile && hasH && hasB && hass && hast))
                    return r;
            }

            return fromRow;
        }

        #region Merge additional sheets (geometry, bolts, weld)

        /// <summary>
        /// Ключевые заголовки для колонки с кодом соединения.
        /// </summary>
        private static readonly string[] KeyColumnHeaders =
            ["CONNECTION_CODE", "Connection_Code", "Code", "Код"];

        /// <summary>
        /// Общая карта отображения колонок листа "geometry" на свойства Row.
        /// </summary>
        private static readonly Dictionary<string, Action<Row, string>> GeometryColumnMap = BuildGeometryColumnMap();

        /// <summary>
        /// Карта отображения колонок листа "weld" на свойства Row.
        /// </summary>
        private static readonly Dictionary<string, Action<Row, string>> WeldColumnMap =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["GOST_weld"] = (r, v) => r.GostWeld = v,
                ["GostWeld"] = (r, v) => r.GostWeld = v,
                ["kf1"] = (r, v) => r.kf1 = v,
                ["kf2"] = (r, v) => r.kf2 = v,
                ["kf3"] = (r, v) => r.kf3 = v,
                ["kf4"] = (r, v) => r.kf4 = v,
                ["kf5"] = (r, v) => r.kf5 = v,
                ["kf6"] = (r, v) => r.kf6 = v,
                ["kf7"] = (r, v) => r.kf7 = v,
                ["kf8"] = (r, v) => r.kf8 = v,
                ["kf9"] = (r, v) => r.kf9 = v,
                ["kf10"] = (r, v) => r.kf10 = v,
                ["kfws"] = (r, v) => r.K_fws_base = v,
                ["Anchor_k_fws_base"] = (r, v) => r.K_fws_base = v
            };

        /// <summary>
        /// Карта отображения колонок листа "bolts" на свойства Row (создаётся динамически).
        /// </summary>
        private static readonly Dictionary<string, Action<Row, string>> BoltsColumnMap = BuildBoltsColumnMap();

        /// <summary>
        /// Карта отображения колонок листа "holes" на свойства Row (создаётся динамически).
        /// </summary>
        private static readonly Dictionary<string, Action<Row, string>> HolesColumnMap = BuildHolesMap();


        private static Dictionary<string, Action<Row, string>> BuildGeometryColumnMap()
        {
            var map = new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase);
            map["GOST"] = (r, v) => r.Gost = v;
            map["Gost"] = (r, v) => r.Gost = v;
            map["GostColumn"] = (r, v) => r.GostColumn = v;
            map["GostBeams"] = (r, v) => r.GostBeams = v;
            map["Марка опорного столика"] = (r, v) => r.TableBrand = v;
            map["Маркаопорногостолика"] = (r, v) => r.TableBrand = v;
            //Фланец
            map["H"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_H = value;
                if (r.H_Plate == 0)
                    r.H_Plate = value;
            };
            map["H_flange"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_H = value;
                if (r.H_Plate == 0)
                    r.H_Plate = value;
            };
            map["B"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_B = value;
                if (r.B_Plate == 0)
                    r.B_Plate = value;
            };
            map["B_flange"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_B = value;
                if (r.B_Plate == 0)
                    r.B_Plate = value;
            };
            map["tp"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_t = value;
                if (r.Tp_Plate == 0)
                    r.Tp_Plate = value;
            };
            map["Tp_flange"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Flange_t = value;
                if (r.Tp_Plate == 0)
                    r.Tp_Plate = value;
            };
            map["Lb"] = (r, v) => r.Flange_Lb = NumericParser.ParseDouble(v);
            map["Lb_flange"] = (r, v) => r.Flange_Lb = NumericParser.ParseDouble(v);
            //Пластина
            map["B_plate"] = (r, v) => r.B_Plate = NumericParser.ParseDouble(v);
            map["H_plate"] = (r, v) => r.H_Plate = NumericParser.ParseDouble(v);
            map["Lws_plate"] = (r, v) => r.Lws_Plate = NumericParser.ParseDouble(v);
            map["Tp_plate"] = (r, v) => r.Tp_Plate = NumericParser.ParseDouble(v);
            map["Tr1_plate"] = (r, v) => r.Tr1_Plate = NumericParser.ParseDouble(v);
            map["Tr2_plate"] = (r, v) => r.Tr2_Plate = NumericParser.ParseDouble(v);
            //Ребра жесткости
            map["B_stiff"] = (r, v) => r.B_Stiff = NumericParser.ParseDouble(v);
            map["H_stiff"] = (r, v) => r.H_Stiff = NumericParser.ParseDouble(v);
            map["Hh"] = (r, v) => r.Hh_Stiff = NumericParser.ParseDouble(v);
            map["Tg_stiff"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Tg_Stiff = value;
                if (r.Tg_Stiff == 0)
                    r.Tg_Stiff = value;
            };
            map["tg"] = (r, v) => r.Tg_Stiff = NumericParser.ParseDouble(v);
            map["Lg_stiff"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Lg_Stiff = value;
                if (r.Lg_Stiff == 0)
                    r.Lg_Stiff = value;
            };
            map["Lg"] = (r, v) => r.Lg_Stiff = NumericParser.ParseDouble(v);
            map["Tf_stiff"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Tf_Stiff = value;
                if (r.Tf_Stiff == 0)
                    r.Tf_Stiff = value;
            };
            map["tf"] = (r, v) => r.Tf_Stiff = NumericParser.ParseDouble(v);
            map["Lh_stiff"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Lh_Stiff = value;
                if (r.Lh_Stiff == 0)
                    r.Lh_Stiff = value;
            };
            map["Lh"] = (r, v) => r.Lh_Stiff = NumericParser.ParseDouble(v);
            map["Hh_stiff"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Hh_Stiff = value;
                if (r.Hh_Stiff == 0)
                    r.Hh_Stiff = value;
            };
            map["Lws_stiff"] = (r, v) => r.Lws_Stiff = NumericParser.ParseDouble(v);
            map["Tp_stiff"] = (r, v) => r.Tp_Stiff = NumericParser.ParseDouble(v);
            map["Tr1_stiff"] = (r, v) => r.Tr1_Stiff = NumericParser.ParseDouble(v);
            map["Tr2_stiff"] = (r, v) => r.Tr2_Stiff = NumericParser.ParseDouble(v);
            map["tr1"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Tr1_Stiff = value;
                if (r.Tr1_Plate == 0)
                    r.Tr1_Plate = value;
            };
            map["tr2"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Tr2_Stiff = value;
                if (r.Tr2_Plate == 0)
                    r.Tr2_Plate = value;
            };
            map["Lst"] = (r, v) =>
            {
                if (r.H_Stiff == 0)
                    r.H_Stiff = NumericParser.ParseDouble(v);
            };
            map["tbp"] = (r, v) =>
            {
                var value = NumericParser.ParseDouble(v);
                r.Tp_Stiff = value;
            };
            map["F_base"] = (r, v) => r.F_base = NumericParser.ParseDouble(v);
            map["Anchor_F_base"] = (r, v) => r.F_base = NumericParser.ParseDouble(v);
            map["Lp_base"] = (r, v) => r.Lp_base = NumericParser.ParseDouble(v);
            map["Anchor_Lp_base"] = (r, v) => r.Lp_base = NumericParser.ParseDouble(v);
            map["Ls_base"] = (r, v) => r.Ls_base = NumericParser.ParseDouble(v);
            map["Anchor_Ls_base"] = (r, v) => r.Ls_base = NumericParser.ParseDouble(v);
            map["Lws"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v);
            map["Lws_base"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v);
            map["Anchor_Lws"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v);
            map["tws"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v);
            map["Tws_base"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v);
            map["Anchor_tws_base"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v);
            map["Dws"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v);
            map["D_ws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v);
            map["Anchor_d_ws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v);
            map["Dp"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v);
            map["D_p_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v);
            map["Anchor_d_p_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v);
            map["xh"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v);
            map["Xh_base"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v);
            map["Anchor_xh_base"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v);
            map["Nh_1_2"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v);
            map["Nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v);
            map["Nh1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v);
            map["Anchor_nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v);
            map["Nh_3_4"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v);
            map["Nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v);
            map["Nh2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v);
            map["Anchor_nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v);
            map["Anchor_var_1"] = (r, v) => r.Anchor_var_1 = v;
            map["Anchor_var_2"] = (r, v) => r.Anchor_var_2 = v;
            map["Anchor_var_3"] = (r, v) => r.Anchor_var_3 = v;
            map["Anchor_var_4"] = (r, v) => r.Anchor_var_4 = v;
            map["Anchor_anchor_var_1"] = (r, v) => r.Anchor_var_1 = v;
            map["Anchor_anchor_var_2"] = (r, v) => r.Anchor_var_2 = v;
            map["Anchor_anchor_var_3"] = (r, v) => r.Anchor_var_3 = v;
            map["Anchor_anchor_var_4"] = (r, v) => r.Anchor_var_4 = v;

            return map;
        }

        /// <summary>
        /// Создаёт карту отображения колонок листа "bolts" на свойства Row.
        /// Включает варианты написания для "Марка опорного столика".
        /// </summary>
        /// <returns>Словарь с отображением заголовков на действия обновления Row.</returns>
        private static Dictionary<string, Action<Row, string>> BuildBoltsColumnMap()
        {
            var map = new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
            {
                ["Option"] = (r, v) => r.OptionBolts = NumericParser.ParseInt(v),
                ["GOST_anchor"] = (r, v) => r.GostAnchore = v,
                ["GOST_anchors"] = (r, v) => r.GostAnchore = v,
                ["GOST_bolts"] = (r, v) => r.GostBolts = v,
                ["GostAnchore"] = (r, v) => r.GostAnchore = v,
                ["GostBolts"] = (r, v) => r.GostBolts = v,
                ["TypeNode"] = (r, v) => r.TypeNode = v,
                ["Тип узла"] = (r, v) => r.TypeNode = v,
                ["Вид узла"] = (r, v) => r.TypeNode = v,
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
                ["p2"] = (r, v) => r.p2 = NumericParser.ParseInt(v),
                ["p3"] = (r, v) => r.p3 = NumericParser.ParseInt(v),
                ["p4"] = (r, v) => r.p4 = NumericParser.ParseInt(v),
                ["p5"] = (r, v) => r.p5 = NumericParser.ParseInt(v),
                ["p6"] = (r, v) => r.p6 = NumericParser.ParseInt(v),
                ["p7"] = (r, v) => r.p7 = NumericParser.ParseInt(v),
                ["p8"] = (r, v) => r.p8 = NumericParser.ParseInt(v),
                ["p9"] = (r, v) => r.p9 = NumericParser.ParseInt(v),
                ["p10"] = (r, v) => r.p10 = NumericParser.ParseInt(v),
                ["Марка опорного столика"] = (r, v) => r.TableBrand = v,
                ["Маркаопорногостолика"] = (r, v) => r.TableBrand = v,
                ["марка"] = (r, v) => r.TableBrand = v,
                ["Марка"] = (r, v) => r.TableBrand = v,
                ["Lp_base"] = (r, v) => r.Lp_base = NumericParser.ParseDouble(v),
                ["Ls_base"] = (r, v) => r.Ls_base = NumericParser.ParseDouble(v),
                ["Анкер1"] = (r, v) => r.Anchor_var_1 = v,
                ["Anchor1"] = (r, v) => r.Anchor_var_1 = v,
                ["Анкер2"] = (r, v) => r.Anchor_var_2 = v,
                ["Anchor2"] = (r, v) => r.Anchor_var_2 = v,
                ["Анкер3"] = (r, v) => r.Anchor_var_3 = v,
                ["Anchor3"] = (r, v) => r.Anchor_var_3 = v,
                ["Анкер4"] = (r, v) => r.Anchor_var_4 = v,
                ["Anchor4"] = (r, v) => r.Anchor_var_4 = v,
                ["Anchor_Lp_base"] = (r, v) => r.Lp_base = NumericParser.ParseDouble(v),
                ["Anchor_Ls_base"] = (r, v) => r.Ls_base = NumericParser.ParseDouble(v),
                ["Anchor_tws_base"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v),
                ["Anchor_d_ws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v),
                ["Anchor_Lws"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v),
                ["Anchor_d_p_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v),
                ["Anchor_xh_base"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v),
                ["Anchor_xh_holes"] = (r, v) => r.Anchor_xh_holes = NumericParser.ParseDouble(v),
                ["Anchor_nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v),
                ["Anchor_nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v),
                ["Anchor_anchor_var_1"] = (r, v) => r.Anchor_var_1 = v,
                ["Anchor_anchor_var_2"] = (r, v) => r.Anchor_var_2 = v,
                ["Anchor_anchor_var_3"] = (r, v) => r.Anchor_var_3 = v,
                ["Anchor_anchor_var_4"] = (r, v) => r.Anchor_var_4 = v,
                ["Anchor_var_1"] = (r, v) => r.Anchor_var_1 = v,
                ["Anchor_var_2"] = (r, v) => r.Anchor_var_2 = v,
                ["Anchor_var_3"] = (r, v) => r.Anchor_var_3 = v,
                ["Anchor_var_4"] = (r, v) => r.Anchor_var_4 = v,
                ["Lws"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v),
                ["Lws_base"] = (r, v) => r.Lws_base = NumericParser.ParseDouble(v),
                ["tws"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v),
                ["Tws_base"] = (r, v) => r.Tws_base = NumericParser.ParseDouble(v),
                ["Dws"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v),
                ["D_ws_base"] = (r, v) => r.D_ws_base = NumericParser.ParseDouble(v),
                ["Dp"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v),
                ["D_p_base"] = (r, v) => r.D_p_base = NumericParser.ParseDouble(v),
                ["xh"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v),
                ["Xh_base"] = (r, v) => r.Xh_base = NumericParser.ParseDouble(v),
                ["Nh_1_2"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v),
                ["Nh_base_var1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v),
                ["Nh1"] = (r, v) => r.Nh_base_var1 = NumericParser.ParseDouble(v),
                ["Nh_3_4"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v),
                ["Nh_base_var2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v),
                ["Nh2"] = (r, v) => r.Nh_base_var2 = NumericParser.ParseDouble(v),
                ["B_plate"] = (r, v) => r.B_Plate = NumericParser.ParseDouble(v),
                ["H_plate"] = (r, v) => r.H_Plate = NumericParser.ParseDouble(v),
                ["Lws_plate"] = (r, v) => r.Lws_Plate = NumericParser.ParseDouble(v),
                ["Tp_plate"] = (r, v) => r.Tp_Plate = NumericParser.ParseDouble(v),
                ["Tr1_plate"] = (r, v) => r.Tr1_Plate = NumericParser.ParseDouble(v),
                ["Tr2_plate"] = (r, v) => r.Tr2_Plate = NumericParser.ParseDouble(v),
                ["B_stiff"] = (r, v) => r.B_Stiff = NumericParser.ParseDouble(v),
                ["H_stiff"] = (r, v) => r.H_Stiff = NumericParser.ParseDouble(v),
                ["Lws_stiff"] = (r, v) => r.Lws_Stiff = NumericParser.ParseDouble(v),
                ["Tp_stiff"] = (r, v) => r.Tp_Stiff = NumericParser.ParseDouble(v),
                ["Tr1_stiff"] = (r, v) => r.Tr1_Stiff = NumericParser.ParseDouble(v),
                ["Tr2_stiff"] = (r, v) => r.Tr2_Stiff = NumericParser.ParseDouble(v)
            };
            return map;
        }

        /// <summary>
        /// Создаёт карту отображения колонок листа "bolts" на свойства Row.
        /// Включает варианты написания для "Марка опорного столика".
        /// </summary>
        /// <returns>Словарь с отображением заголовков на действия обновления Row.</returns>
        private static Dictionary<string, Action<Row, string>> BuildHolesMap()
        {
            var map = new Dictionary<string, Action<Row, string>>(StringComparer.OrdinalIgnoreCase)
            {
                ["Option"] = (r, v) => r.OptionHoles = NumericParser.ParseInt(v),
                ["GOST"] = (r, v) => r.GostHoles = v,
                ["GostHoles"] = (r, v) => r.GostHoles = v,
                //Радиус отверстия
                ["F"] = (r, v) => { r.F_holes = NumericParser.ParseInt(v); r.N_Rows = 1; },
                ["F_holes"] = (r, v) => r.F_holes = NumericParser.ParseInt(v),
                ["DiameterHolesForBolts"] = (r, v) => r.F_holes = NumericParser.ParseInt(v),
                //Марка опорного столика
                ["Марка опорного столика"] = (r, v) => r.TableBrandHoles = v,
                ["Маркаопорногостолика"] = (r, v) => r.TableBrandHoles = v,
                ["марка"] = (r, v) => r.TableBrandHoles = v,
                ["Марка"] = (r, v) => r.TableBrandHoles = v,
                //Гост для анкеров
                ["GOST_holes"] = (r, v) => r.GostHoles = v,
                //Радиус отверстия под анкер
                ["Dws_holes"] = (r, v) => r.Dws_holes = NumericParser.ParseDouble(v),
                //Радиус отверстия под анкер
                ["Dp_holes"] = (r, v) => r.Dp_holes = NumericParser.ParseDouble(v),
                //Количество отверстий от 1 до 4
                ["Nh_holes_1_4"] = (r, v) => r.Nh_Holes_1_4 = NumericParser.ParseInt(v),
                //Количество отверстий от 1 до 8
                ["Nh_holes_5_8"] = (r, v) => r.Nh_Holes_5_8 = NumericParser.ParseInt(v),
                ["Nh_holes_1_8"] = (r, v) => r.Nh_Holes_5_8 = NumericParser.ParseInt(v),
                //Расстояние между отверстиями
                ["Anchor_xh_holes"] = (r, v) => r.Anchor_xh_holes = NumericParser.ParseDouble(v),
                ["xh"] = (r, v) => r.Anchor_xh_holes = NumericParser.ParseDouble(v),

            };
            return map;
        }

        /// <summary>
        /// Гарантирует, что список CoordinatesBolts содержит достаточное количество элементов.
        /// </summary>
        /// <param name="r">Объект Row.</param>
        /// <param name="count">Требуемое количество элементов.</param>
        private static void EnsureBolts(Row r, int count)
        {
            while (r.CoordinatesBolts.Count < count)
                r.CoordinatesBolts.Add(new CoordinatesBolts(0, 0, 0));
        }

        /// <summary>
        /// Объединяет данные из дополнительных листов (geometry, bolts, weld) с основным списком Row.
        /// </summary>
        /// <param name="package">Пакет Excel.</param>
        /// <param name="mainWs">Основной лист.</param>
        /// <param name="list">Список объектов Row для обогащения.</param>
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
                if (string.Equals(sheetName, "data", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (string.Equals(sheetName, "geometry", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, GeometryColumnMap, codeLookup, list);
                else if (string.Equals(sheetName, "bolts", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, BoltsColumnMap, codeLookup, list);
                else if (string.Equals(sheetName, "holes", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, HolesColumnMap, codeLookup, list);
                else if (string.Equals(sheetName, "weld", StringComparison.OrdinalIgnoreCase))
                    MergeSheet(ws, WeldColumnMap, codeLookup, list);
            }
        }

        /// <summary>
        /// Объединяет данные из одного дополнительного листа с основным списком Row.
        /// Использует CONNECTION_CODE или позицию строки для сопоставления.
        /// </summary>
        /// <param name="ws">Дополнительный лист Excel.</param>
        /// <param name="propertyMap">Карта отображения заголовков на свойства Row.</param>
        /// <param name="codeLookup">Словарь для поиска Row по CONNECTION_CODE.</param>
        /// <param name="list">Список объектов Row.</param>
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

            int headerRow = FindSheetHeaderRow(ws, startRow, Math.Min(endRow, startRow + 30), startCol, endCol, propertyMap);

            var headers = new List<string>();
            for (int c = startCol; c <= endCol; c++)
                headers.Add(HeaderUtils.NormalizeHeader((ws.Cells[headerRow, c].Text ?? "").Trim()));

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

            for (int r = headerRow + 1; r <= endRow; r++)
            {
                Row? target = null;

                string? key = null;
                if (keyCol >= 0)
                {
                    key = (ws.Cells[r, startCol + keyCol].Text ?? "").Trim();
                    if (!string.IsNullOrWhiteSpace(key))
                        codeLookup.TryGetValue(key, out target);
                }

                // Стратегия 2: Индексный поиск (строки обычно совпадают по порядку)
                if (target == null)
                {
                    int idx = r - headerRow - 1;
                    if (idx >= 0 && idx < list.Count)
                        target = list[idx];
                }

                // Стратегия 3: Если обе стратегии не сработали, но данные есть, 
                // попробуем найти первую незаполненную строку с TableBrand=="" для bolts листа
                if (target == null)
                {
                    // Проверяем, есть ли вообще данные в этой строке
                    bool hasData = false;
                    foreach (var (colIdx, _) in colMappings)
                    {
                        var text = (ws.Cells[r, startCol + colIdx].Text ?? "").Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            hasData = true;
                            break;
                        }
                    }

                    if (hasData)
                    {
                        // Для листа bolts: найдем первую строку с пустым TableBrand
                        foreach (var row in list)
                        {
                            if (string.IsNullOrEmpty(row.TableBrand))
                            {
                                target = row;
                                break;
                            }
                        }
                    }
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

        private static int FindSheetHeaderRow(
            ExcelWorksheet ws,
            int fromRow,
            int toRow,
            int startCol,
            int endCol,
            IReadOnlyDictionary<string, Action<Row, string>> propertyMap)
        {
            for (int r = fromRow; r <= toRow; r++)
            {
                var tokens = new List<string>();
                for (int c = startCol; c <= endCol; c++)
                    tokens.Add(HeaderUtils.NormalizeHeader((ws.Cells[r, c].Text ?? "").Trim()));

                bool hasKey = HeaderUtils.IndexOfHeaderAny(tokens, KeyColumnHeaders) >= 0;
                int mappedCount = tokens.Count(t => !string.IsNullOrWhiteSpace(t) && propertyMap.ContainsKey(t));

                if (mappedCount >= 2 || (hasKey && mappedCount >= 1))
                    return r;
            }

            return fromRow;
        }
        #endregion
    }
}