using ConvertData.Domain;
using ConvertData.Infrastructure.Interop;
using OfficeOpenXml;

namespace ConvertData.Infrastructure;

/// <summary>
/// Вспомогательные методы для чтения и обработки листов Excel через EPPlus.
/// </summary>
internal static class EpplusWorksheetHelpers
{
    /// <summary>
    /// Возвращает границы заполненной области листа.
    /// </summary>
    /// <param name="worksheet">Лист Excel.</param>
    /// <returns>Границы заполненной области.</returns>
    internal static EpplusWorksheetBounds GetBounds(ExcelWorksheet worksheet)
    {
        return new EpplusWorksheetBounds(
            worksheet.Dimension.Start.Row,
            worksheet.Dimension.End.Row,
            worksheet.Dimension.Start.Column,
            worksheet.Dimension.End.Column);
    }

    /// <summary>
    /// Возвращает текст ячейки по указанным координатам.
    /// </summary>
    /// <param name="worksheet">Лист Excel.</param>
    /// <param name="row">Номер строки.</param>
    /// <param name="col">Номер столбца.</param>
    /// <returns>Обрезанное текстовое значение ячейки или пустую строку.</returns>
    internal static string GetCell(ExcelWorksheet worksheet, int row, int? col)
    {
        if (col == null)
            return "";

        return (worksheet.Cells[row, col.Value].Text ?? "").Trim();
    }

    /// <summary>
    /// Проверяет, присутствует ли в карте колонок хотя бы один профильный столбец.
    /// </summary>
    /// <param name="map">Карта колонок Excel.</param>
    /// <returns><see langword="true"/>, если найден хотя бы один профильный столбец.</returns>
    internal static bool HasAnyProfileColumns(ExcelColumnMap map)
    {
        return map.IdxProfileBeam >= 0
            || map.IdxProfileColumn >= 0
            || map.IdxProfileBrace >= 0
            || map.IdxProfileRigel >= 0
            || map.IdxProfileRunThrough >= 0;
    }

    /// <summary>
    /// Применяет сопоставленные колонки листа к целевой строке доменной модели.
    /// </summary>
    /// <param name="row">Целевая строка.</param>
    /// <param name="worksheet">Лист Excel.</param>
    /// <param name="rowIndex">Индекс строки на листе.</param>
    /// <param name="startCol">Начальный столбец диапазона.</param>
    /// <param name="headers">Нормализованные заголовки.</param>
    /// <param name="propertyMap">Карта сопоставления заголовков со свойствами строки.</param>
    internal static void ApplyMappedColumns(
        Row row,
        ExcelWorksheet worksheet,
        int rowIndex,
        int startCol,
        IReadOnlyList<string> headers,
        IReadOnlyDictionary<string, Action<Row, string>> propertyMap,
        ISet<string>? excludedHeaders = null)
    {
        for (int i = 0; i < headers.Count; i++)
        {
            var headerName = headers[i];
            if (string.IsNullOrWhiteSpace(headerName)
                || excludedHeaders?.Contains(headerName) == true
                || !propertyMap.TryGetValue(headerName, out var setter))
                continue;

            var value = GetCell(worksheet, rowIndex, startCol + i);
            if (!string.IsNullOrWhiteSpace(value))
                setter(row, value);
        }
    }

    /// <summary>
    /// Гарантирует наличие заданного количества координат болтов в строке.
    /// </summary>
    /// <param name="row">Целевая строка.</param>
    /// <param name="count">Минимальное количество координат.</param>
    internal static void EnsureBolts(Row row, int count)
    {
        while (row.CoordinatesBolts.Count < count)
            row.CoordinatesBolts.Add(new CoordinatesBolts(0, 0, 0));
    }
}
