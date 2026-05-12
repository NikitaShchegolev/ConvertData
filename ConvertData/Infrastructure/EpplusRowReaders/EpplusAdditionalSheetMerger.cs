using ConvertData.Domain;
using ConvertData.Infrastructure.Parsing;

using OfficeOpenXml;

namespace ConvertData.Infrastructure;

/// <summary>
/// Объединяет данные дополнительных листов книги с основной коллекцией строк.
/// </summary>
internal sealed class EpplusAdditionalSheetMerger
{
    /// <summary>
    /// Выполняет объединение дополнительных листов книги с уже считанными строками.
    /// </summary>
    /// <param name="package">Открытый пакет Excel.</param>
    /// <param name="mainWorksheet">Основной лист с данными.</param>
    /// <param name="rows">Коллекция строк, в которую нужно влить дополнительные данные.</param>
    public void Merge(ExcelPackage package, ExcelWorksheet mainWorksheet, List<Row> rows)
    {
        if (rows.Count == 0)
            return;

        var codeLookup = BuildCodeLookup(rows);
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            if (worksheet == mainWorksheet || worksheet.Dimension == null)
                continue;

            var sheetName = (worksheet.Name ?? "").Trim();
            if (string.Equals(sheetName, "data", StringComparison.OrdinalIgnoreCase))
                continue;

            if (!EpplusRowPropertyMaps.TryGetSheetMap(sheetName, out var propertyMap))
                continue;

            MergeSheet(worksheet, propertyMap, codeLookup, rows);
        }
    }

    /// <summary>
    /// Строит словарь строк по коду соединения.
    /// </summary>
    private static Dictionary<string, Row> BuildCodeLookup(List<Row> rows)
    {
        var codeLookup = new Dictionary<string, Row>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            if (!string.IsNullOrWhiteSpace(row.CONNECTION_CODE) && !codeLookup.ContainsKey(row.CONNECTION_CODE))
                codeLookup[row.CONNECTION_CODE] = row;
        }

        return codeLookup;
    }

    /// <summary>
    /// Объединяет данные одного дополнительного листа.
    /// </summary>
    private static void MergeSheet(
        ExcelWorksheet worksheet,
        Dictionary<string, Action<Row, string>> propertyMap,
        Dictionary<string, Row> codeLookup,
        List<Row> rows)
    {
        var bounds = EpplusWorksheetHelpers.GetBounds(worksheet);
        int headerRow = FindHeaderRow(worksheet, bounds, propertyMap);
        var headers = ReadHeaders(worksheet, headerRow, bounds.StartCol, bounds.EndCol);
        int keyCol = FindKeyColumn(headers);
        var mappings = BuildColumnMappings(headers, keyCol, propertyMap);
        if (mappings.Count == 0)
            return;

        int dataRowOrdinal = 0;
        for (int r = headerRow + 1; r <= bounds.EndRow; r++)
        {
            if (!HasAnyData(worksheet, r, bounds.StartCol, mappings))
                continue;

            var target = ResolveTargetRow(worksheet, r, bounds.StartCol, keyCol, mappings, codeLookup, rows, dataRowOrdinal);
            if (target == null)
            {
                dataRowOrdinal++;
                continue;
            }

            ApplyMappings(worksheet, r, bounds.StartCol, mappings, target);
            dataRowOrdinal++;
        }
    }

    /// <summary>
    /// Находит строку заголовков на дополнительном листе.
    /// </summary>
    private static int FindHeaderRow(
        ExcelWorksheet worksheet,
        EpplusWorksheetBounds bounds,
        IReadOnlyDictionary<string, Action<Row, string>> propertyMap)
    {
        for (int r = bounds.StartRow; r <= Math.Min(bounds.EndRow, bounds.StartRow + 30); r++)
        {
            var tokens = ReadHeaders(worksheet, r, bounds.StartCol, bounds.EndCol);
            bool hasKey = HeaderUtils.IndexOfHeaderAny(tokens, EpplusRowPropertyMaps.KeyColumnHeaders) >= 0;
            int mappedCount = tokens.Count(t => !string.IsNullOrWhiteSpace(t) && propertyMap.ContainsKey(t));
            if (mappedCount >= 2 || (hasKey && mappedCount >= 1))
                return r;
        }

        return bounds.StartRow;
    }

    /// <summary>
    /// Считывает и нормализует заголовки из указанной строки листа.
    /// </summary>
    private static List<string> ReadHeaders(ExcelWorksheet worksheet, int row, int startCol, int endCol)
    {
        var headers = new List<string>();
        for (int c = startCol; c <= endCol; c++)
            headers.Add(HeaderUtils.NormalizeHeader((worksheet.Cells[row, c].Text ?? "").Trim()));

        return headers;
    }

    /// <summary>
    /// Находит индекс ключевого столбца по известным названиям.
    /// </summary>
    private static int FindKeyColumn(IReadOnlyList<string> headers)
    {
        for (int i = 0; i < headers.Count; i++)
        {
            foreach (var name in EpplusRowPropertyMaps.KeyColumnHeaders)
            {
                if (string.Equals(headers[i], name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Строит список сопоставлений между индексами колонок и сеттерами свойств.
    /// </summary>
    private static List<(int ColumnIndex, Action<Row, string> Setter)> BuildColumnMappings(
        IReadOnlyList<string> headers,
        int keyCol,
        IReadOnlyDictionary<string, Action<Row, string>> propertyMap)
    {
        var mappings = new List<(int ColumnIndex, Action<Row, string> Setter)>();
        for (int i = 0; i < headers.Count; i++)
        {
            if (i == keyCol || string.IsNullOrWhiteSpace(headers[i]))
                continue;

            if (propertyMap.TryGetValue(headers[i], out var setter))
                mappings.Add((i, setter));
        }

        return mappings;
    }

    /// <summary>
    /// Определяет целевую строку для применения данных дополнительного листа.
    /// </summary>
    private static Row? ResolveTargetRow(
        ExcelWorksheet worksheet,
        int rowIndex,
        int startCol,
        int keyCol,
        IReadOnlyList<(int ColumnIndex, Action<Row, string> Setter)> mappings,
        IReadOnlyDictionary<string, Row> codeLookup,
        IReadOnlyList<Row> rows,
        int dataRowOrdinal)
    {
        var target = ResolveByConnectionCode(worksheet, rowIndex, startCol, keyCol, codeLookup);
        if (target != null)
            return target;

        target = ResolveByDataRowOrdinal(rows, dataRowOrdinal);
        if (target != null)
            return target;

        return null;
    }

    /// <summary>
    /// Пытается найти строку по коду соединения.
    /// </summary>
    private static Row? ResolveByConnectionCode(
        ExcelWorksheet worksheet,
        int rowIndex,
        int startCol,
        int keyCol,
        IReadOnlyDictionary<string, Row> codeLookup)
    {
        if (keyCol < 0)
            return null;

        var key = (worksheet.Cells[rowIndex, startCol + keyCol].Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(key))
            return null;

        codeLookup.TryGetValue(key, out var target);
        return target;
    }
    /// <summary>
    /// Пытается сопоставить строку по порядковому индексу непустой строки данных.
    /// </summary>
    private static Row? ResolveByDataRowOrdinal(IReadOnlyList<Row> rows, int dataRowOrdinal)
    {
        if (dataRowOrdinal >= 0 && dataRowOrdinal < rows.Count)
            return rows[dataRowOrdinal];

        return null;
    }

    /// <summary>
    /// Проверяет, содержит ли строка хотя бы одно значение из сопоставленных колонок.
    /// </summary>
    private static bool HasAnyData(
        ExcelWorksheet worksheet,
        int rowIndex,
        int startCol,
        IReadOnlyList<(int ColumnIndex, Action<Row, string> Setter)> mappings)
    {
        foreach (var (columnIndex, _) in mappings)
        {
            var text = (worksheet.Cells[rowIndex, startCol + columnIndex].Text ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(text))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Применяет значения сопоставленных колонок к целевой строке.
    /// </summary>
    private static void ApplyMappings( ExcelWorksheet worksheet, int rowIndex, int startCol, IReadOnlyList<(int ColumnIndex, Action<Row, string> Setter)> mappings, Row target)
    {
        foreach (var (columnIndex, setter) in mappings)
        {
            var text = (worksheet.Cells[rowIndex, startCol + columnIndex].Text ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(text))
                setter(target, text);
        }
    }
}
