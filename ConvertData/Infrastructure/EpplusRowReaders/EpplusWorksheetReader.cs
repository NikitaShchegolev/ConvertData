using ConvertData.Domain;
using ConvertData.Enums;
using ConvertData.Infrastructure.Interop;
using ConvertData.Infrastructure.Parsing;

using OfficeOpenXml;

namespace ConvertData.Infrastructure;

/// <summary>
/// Читает отдельный лист Excel и преобразует его в коллекцию доменных строк.
/// </summary>
internal sealed class EpplusWorksheetReader
{
    private static readonly ProfileSectionType[] ProfileSections =
    [
        ProfileSectionType.Beam,
        ProfileSectionType.Column,
        ProfileSectionType.Brace,
        ProfileSectionType.Rigel,
        ProfileSectionType.RunThrougth
    ];

    private static readonly ProfileSectionType[] MainRowMappingTypes =
    [
        ProfileSectionType.Stiffness,
        ProfileSectionType.Forces,
        ProfileSectionType.Coefficients,
        ProfileSectionType.BeamGeometry,
        ProfileSectionType.PlateGeometry,
        ProfileSectionType.StiffGeometry,
        ProfileSectionType.Base,
        ProfileSectionType.Anchor,
        ProfileSectionType.ShearKey,
        ProfileSectionType.Brace
    ];

    private static readonly HashSet<string> MainRowMappedGeometryHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Марка опорного столика",
        "ProfileBeam",
        "ProfileColumn",
        "ProfileBrace",
        "ProfileRigel",
        "ProfileRunThrough",
        "ProfileRunThrougth",
        "ProfileRunTrought",
        "GostBeams",
        "GostColumn",
        "Beam_H",
        "Beam_B",
        "Beam_s",
        "Beam_t",
        "Lb_plate",
        "B_plate",
        "H_plate",
        "Lws_plate",
        "Tp_plate",
        "Tr1_plate",
        "Tr2_plate",
        "B_stiff",
        "H_stiff",
        "Lws_stiff",
        "Tp_stiff",
        "Tr1_stiff",
        "Tr2_stiff",
        "Tg_stiff",
        "Lg_stiff",
        "Tf_stiff",
        "Twp_stiff",
        "Lh_stiff",
        "Hh_stiff",
        "F_base",
        "H_base",
        "B_base",
        "S_base",
        "Lp_base",
        "Ls_base",
        "Lws_base",
        "T_base",
        "Tws_base",
        "Dws_base",
        "D_ws_base",
        "Dp_base",
        "D_p_base",
        "xh_base",
        "K_fws_base",
        "k_fws_base",
        "Nh_base_var1",
        "Nh_base_var2",
        "Anchor_var_1",
        "Anchor_var_2",
        "Anchor_var_3",
        "Anchor_var_4",
        "GostBolts",
        "Lp_shearKey",
        "Ls_shearKey",
        "Lb_brace",
        "a_brace",
        "e2_brace",
        "e3_brace",
        "n1_brace",
        "n2_brace"
    };

    private static readonly HashSet<string> ProfileRowMappedGeometryHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "ProfileBeam",
        "ProfileColumn",
        "ProfileBrace",
        "ProfileRigel",
        "ProfileRunThrough",
        "ProfileRunThrougth",
        "ProfileRunTrought",
        "GostBolts",
        "GostBeams",
        "GostColumn",
        "GostBrace",
        "GostRigel",
        "GostRunThrougth",
        "GostRunThrougth",
        "GostRunTrought",
        "Lp_shearKey",
        "Ls_shearKey",
        "Beam_H",
        "Beam_B",
        "Beam_s",
        "Beam_t",
        "Column_H",
        "Column_B",
        "Column_s",
        "Column_t",
        "Brace_H",
        "Brace_B",
        "Brace_s",
        "Brace_t",
        "Rigel_H",
        "Rigel_B",
        "Rigel_s",
        "Rigel_t",
        "RunThrougth_H",
        "RunThrougth_B",
        "RunThrougth_s",
        "RunThrougth_t",
        "RunThrougth_H",
        "RunThrougth_B",
        "RunThrougth_s",
        "RunThrougth_t",
        "RunTrought_H",
        "RunTrought_B",
        "RunTrought_s",
        "RunTrought_t"
    };

    /// <summary>
    /// Считывает строки из листа Excel.
    /// </summary>
    /// <param name="worksheet">Лист Excel.</param>
    /// <returns>Результат чтения листа.</returns>
    public EpplusWorksheetReadResult Read(ExcelWorksheet worksheet)
    {
        var bounds = EpplusWorksheetHelpers.GetBounds(worksheet);
        int headerRow = FindHeaderRow(worksheet, bounds);
        var headers = ReadHeaders(worksheet, headerRow, bounds.StartCol, bounds.EndCol);
        var map = ExcelHeaderResolver.Resolve(headers);

        PrepareColumnMap(map);

        if (!map.IsMainTable && !EpplusWorksheetHelpers.HasAnyProfileColumns(map))
            throw new InvalidDataException("Cannot find required headers in worksheet");

        var rows = map.IsMainTable
            ? ReadMainTableRows(worksheet, bounds, headerRow, headers, map)
            : ReadProfileTableRows(worksheet, bounds, headerRow, headers, map);

        return new EpplusWorksheetReadResult(rows, map.IsMainTable);
    }

    /// <summary>
    /// Находит строку заголовков на листе.
    /// </summary>
    private static int FindHeaderRow(ExcelWorksheet worksheet, EpplusWorksheetBounds bounds)
    {
        int headerRow = bounds.StartRow;
        for (int r = bounds.StartRow; r <= Math.Min(bounds.EndRow, bounds.StartRow + 30); r++)
        {
            var tokens = ReadHeaders(worksheet, r, bounds.StartCol, bounds.EndCol);
            var map = ExcelHeaderResolver.Resolve(tokens);
            if (map.IsMainTable || EpplusWorksheetHelpers.HasAnyProfileColumns(map))
            {
                headerRow = r;
                break;
            }
        }

        return headerRow;
    }

    /// <summary>
    /// Считывает и нормализует заголовки из указанной строки.
    /// </summary>
    private static List<string> ReadHeaders(ExcelWorksheet worksheet, int row, int startCol, int endCol)
    {
        var headers = new List<string>();
        for (int c = startCol; c <= endCol; c++)
            headers.Add(HeaderUtils.NormalizeHeader((worksheet.Cells[row, c].Text ?? "").Trim()));

        return headers;
    }

    /// <summary>
    /// Подготавливает карту колонок для чтения таблицы профилей.
    /// </summary>
    private static void PrepareColumnMap(ExcelColumnMap map)
    {
        if (map.IsMainTable)
            return;

        if (map.IdxProfileBeam >= 0
            && (map.IdxH_beam < 0 || map.IdxB_beam < 0 || map.Idxs_beam < 0 || map.Idxt_beam < 0))
        {
            if (map.IdxH_beam < 0) map.IdxH_beam = map.IdxProfileBeam + 1;
            if (map.IdxB_beam < 0) map.IdxB_beam = map.IdxProfileBeam + 2;
            if (map.Idxs_beam < 0) map.Idxs_beam = map.IdxProfileBeam + 3;
            if (map.Idxt_beam < 0) map.Idxt_beam = map.IdxProfileBeam + 4;
            return;
        }

        if (EpplusWorksheetHelpers.HasAnyProfileColumns(map))
            return;

        map.IdxProfileBeam = 0;
        map.IdxH_beam = 1;
        map.IdxB_beam = 2;
        map.Idxs_beam = 3;
        map.Idxt_beam = 4;
    }

    /// <summary>
    /// Считывает строки из основной таблицы данных.
    /// </summary>
    private static List<Row> ReadMainTableRows(
        ExcelWorksheet worksheet,
        EpplusWorksheetBounds bounds,
        int headerRow,
        IReadOnlyList<string> headers,
        ExcelColumnMap map)
    {
        var rows = new List<Row>();
        int firstDataRow = headerRow + 1;

        for (int r = firstDataRow; r <= bounds.EndRow; r++)
        {
            string code = EpplusWorksheetHelpers.GetCell(worksheet, r, bounds.StartCol + map.IdxCode);
            if (string.IsNullOrWhiteSpace(code))
                continue;

            var row = CreateMainRow(worksheet, bounds, map, r, code);

            foreach (var mappingType in MainRowMappingTypes)
                ApplyMainRowMapping(worksheet, bounds, map, r, row, mappingType);

            EpplusWorksheetHelpers.ApplyMappedColumns(row, worksheet, r, bounds.StartCol, headers, EpplusRowPropertyMaps.GeometryColumnMap, MainRowMappedGeometryHeaders);
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>
    /// Считывает строки из профильной таблицы.
    /// </summary>
    private static List<Row> ReadProfileTableRows(
        ExcelWorksheet worksheet,
        EpplusWorksheetBounds bounds,
        int headerRow,
        IReadOnlyList<string> headers,
        ExcelColumnMap map)
    {
        var rows = new List<Row>();
        int firstDataRow = headerRow + 1;

        for (int r = firstDataRow; r <= bounds.EndRow; r++)
        {
            var hasProfileValue = false;
            foreach (var sectionType in ProfileSections)
            {
                int? profileColumn = sectionType switch
                {
                    ProfileSectionType.Beam => bounds.StartCol + map.IdxProfileBeam,
                    ProfileSectionType.Column => map.IdxProfileColumn >= 0 ? bounds.StartCol + map.IdxProfileColumn : null,
                    ProfileSectionType.Brace => map.IdxProfileBrace >= 0 ? bounds.StartCol + map.IdxProfileBrace : null,
                    ProfileSectionType.Rigel => map.IdxProfileRigel >= 0 ? bounds.StartCol + map.IdxProfileRigel : null,
                    ProfileSectionType.RunThrougth => map.IdxProfileRunThrough >= 0 ? bounds.StartCol + map.IdxProfileRunThrough : null,
                    _ => null
                };

                if (string.IsNullOrWhiteSpace(EpplusWorksheetHelpers.GetCell(worksheet, r, profileColumn)))
                    continue;

                hasProfileValue = true;
                break;
            }

            if (!hasProfileValue)
                continue;

            var row = new Row();
            EpplusWorksheetHelpers.ApplyMappedColumns(row, worksheet, r, bounds.StartCol, headers, EpplusRowPropertyMaps.GeometryColumnMap, ProfileRowMappedGeometryHeaders);
            rows.Add(row);
        }

        return rows;
    }

    private static Row CreateMainRow(ExcelWorksheet worksheet, EpplusWorksheetBounds bounds, ExcelColumnMap map, int rowIndex, string code)
    {
        var row = RowMapper.MapMainRowIdentity(
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, bounds.StartCol + map.IdxName),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxStructuralElement >= 0 ? bounds.StartCol + map.IdxStructuralElement : null),
            code,
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTypeNode >= 0 ? bounds.StartCol + map.IdxTypeNode : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGostColumnAndBeams >= 0 ? bounds.StartCol + map.IdxGostColumnAndBeams : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGostBolts >= 0 ? bounds.StartCol + map.IdxGostBolts : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGostAnchore >= 0 ? bounds.StartCol + map.IdxGostAnchore : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGostWeld >= 0 ? bounds.StartCol + map.IdxGostWeld : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGostProfile >= 0 ? bounds.StartCol + map.IdxGostProfile : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTableBrand >= 0 ? bounds.StartCol + map.IdxTableBrand : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, bounds.StartCol + map.IdxProfileBeam),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxProfileColumn >= 0 ? bounds.StartCol + map.IdxProfileColumn : null),
            EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxExplanations >= 0 ? bounds.StartCol + map.IdxExplanations : null));

        row.ProfileBrace = EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxProfileBrace >= 0 ? bounds.StartCol + map.IdxProfileBrace : null);
        row.ProfileRigel = EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxProfileRigel >= 0 ? bounds.StartCol + map.IdxProfileRigel : null);
        row.ProfileRunThrough = EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxProfileRunThrough >= 0 ? bounds.StartCol + map.IdxProfileRunThrough : null);

        return row;
    }

    private static void ApplyMainRowMapping(        ExcelWorksheet worksheet,        EpplusWorksheetBounds bounds,        ExcelColumnMap map,
        int rowIndex,        Row row,        ProfileSectionType mappingType)
    {
        switch (mappingType)
        {
            case ProfileSectionType.Stiffness:
                RowMapper.MapMainRowStiffness(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxVariable >= 0 ? bounds.StartCol + map.IdxVariable : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxSj >= 0 ? bounds.StartCol + map.IdxSj : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxSjo >= 0 ? bounds.StartCol + map.IdxSjo : null));
                break;
            case ProfileSectionType.Forces:
                RowMapper.MapMainRowForces(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxNt >= 0 ? bounds.StartCol + map.IdxNt : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxQy >= 0 ? bounds.StartCol + map.IdxQy : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxQz >= 0 ? bounds.StartCol + map.IdxQz : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxT >= 0 ? bounds.StartCol + map.IdxT : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxNc >= 0 ? bounds.StartCol + map.IdxNc : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxN >= 0 ? bounds.StartCol + map.IdxN : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMy >= 0 ? bounds.StartCol + map.IdxMy : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMy_compression >= 0 ? bounds.StartCol + map.IdxMy_compression : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMy_tension >= 0 ? bounds.StartCol + map.IdxMy_tension : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMneg >= 0 ? bounds.StartCol + map.IdxMneg : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMz >= 0 ? bounds.StartCol + map.IdxMz : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMz_compression >= 0 ? bounds.StartCol + map.IdxMz_compression : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMz_tension >= 0 ? bounds.StartCol + map.IdxMz_tension : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMx >= 0 ? bounds.StartCol + map.IdxMx : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxMw >= 0 ? bounds.StartCol + map.IdxMw : null));
                break;
            case ProfileSectionType.Coefficients:
                RowMapper.MapMainRowCoefficients(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxAlpha >= 0 ? bounds.StartCol + map.IdxAlpha : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxBeta >= 0 ? bounds.StartCol + map.IdxBeta : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxGamma >= 0 ? bounds.StartCol + map.IdxGamma : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxDelta >= 0 ? bounds.StartCol + map.IdxDelta : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxEpsilon >= 0 ? bounds.StartCol + map.IdxEpsilon : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLambda >= 0 ? bounds.StartCol + map.IdxLambda : null));
                break;
            case ProfileSectionType.PlateGeometry:
                RowMapper.MapMainRowPlateGeometry(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLb_plate >= 0 ? bounds.StartCol + map.IdxLb_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxB_plate >= 0 ? bounds.StartCol + map.IdxB_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxH_plate >= 0 ? bounds.StartCol + map.IdxH_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLws_plate >= 0 ? bounds.StartCol + map.IdxLws_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTp_plate >= 0 ? bounds.StartCol + map.IdxTp_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTr1_plate >= 0 ? bounds.StartCol + map.IdxTr1_plate : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTr2_plate >= 0 ? bounds.StartCol + map.IdxTr2_plate : null));
                break;
            case ProfileSectionType.StiffGeometry:
                RowMapper.MapMainRowStiffGeometry(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxB_stiff >= 0 ? bounds.StartCol + map.IdxB_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxH_stiff >= 0 ? bounds.StartCol + map.IdxH_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLws_stiff >= 0 ? bounds.StartCol + map.IdxLws_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idxtp_stiff >= 0 ? bounds.StartCol + map.Idxtp_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idxtr1_stiff >= 0 ? bounds.StartCol + map.Idxtr1_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idxtr2_stiff >= 0 ? bounds.StartCol + map.Idxtr2_stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTg_Stiff >= 0 ? bounds.StartCol + map.IdxTg_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLg_Stiff >= 0 ? bounds.StartCol + map.IdxLg_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTf_Stiff >= 0 ? bounds.StartCol + map.IdxTf_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTwp_Stiff >= 0 ? bounds.StartCol + map.IdxTwp_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxTbp_Stiff >= 0 ? bounds.StartCol + map.IdxTbp_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLh_Stiff >= 0 ? bounds.StartCol + map.IdxLh_Stiff : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxHh_Stiff >= 0 ? bounds.StartCol + map.IdxHh_Stiff : null));
                break;
            case ProfileSectionType.Base:
                RowMapper.MapMainRowBase(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdF_base >= 0 ? bounds.StartCol + map.IdF_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdLws_base >= 0 ? bounds.StartCol + map.IdLws_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdLp_base >= 0 ? bounds.StartCol + map.IdLp_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdLs_base >= 0 ? bounds.StartCol + map.IdLs_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdTws_base >= 0 ? bounds.StartCol + map.IdTws_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdD_ws_base >= 0 ? bounds.StartCol + map.IdD_ws_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdD_p_base >= 0 ? bounds.StartCol + map.IdD_p_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxH_base >= 0 ? bounds.StartCol + map.IdxH_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxS_base >= 0 ? bounds.StartCol + map.IdxS_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxB_base >= 0 ? bounds.StartCol + map.IdxB_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxT_base >= 0 ? bounds.StartCol + map.IdxT_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdK_fws_base >= 0 ? bounds.StartCol + map.IdK_fws_base : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdNh_base_var1 >= 0 ? bounds.StartCol + map.IdNh_base_var1 : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdNh_base_var2 >= 0 ? bounds.StartCol + map.IdNh_base_var2 : null));
                break;
            case ProfileSectionType.Anchor:
                RowMapper.MapMainRowAnchor(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdAnchor_var_1 >= 0 ? bounds.StartCol + map.IdAnchor_var_1 : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdAnchor_var_2 >= 0 ? bounds.StartCol + map.IdAnchor_var_2 : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdAnchor_var_3 >= 0 ? bounds.StartCol + map.IdAnchor_var_3 : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdAnchor_var_4 >= 0 ? bounds.StartCol + map.IdAnchor_var_4 : null));
                break;
            case ProfileSectionType.ShearKey:
                RowMapper.MapMainRowShearKey(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLp_shearKey >= 0 ? bounds.StartCol + map.IdxLp_shearKey : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.IdxLs_shearKey >= 0 ? bounds.StartCol + map.IdxLs_shearKey : null));
                break;
            case ProfileSectionType.Brace:
                RowMapper.MapMainRowBrace(
                    row,
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_a_brace >= 0 ? bounds.StartCol + map.Idx_a_brace : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_e2_brace >= 0 ? bounds.StartCol + map.Idx_e2_brace : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_e3_brace >= 0 ? bounds.StartCol + map.Idx_e3_brace : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_n1_brace >= 0 ? bounds.StartCol + map.Idx_n1_brace : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_n2_brace >= 0 ? bounds.StartCol + map.Idx_n2_brace : null),
                    EpplusWorksheetHelpers.GetCell(worksheet, rowIndex, map.Idx_Lb_brace >= 0 ? bounds.StartCol + map.Idx_Lb_brace : null));
                break;
        }
    }
}
