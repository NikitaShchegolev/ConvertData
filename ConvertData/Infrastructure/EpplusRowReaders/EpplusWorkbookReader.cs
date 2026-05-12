using ConvertData.Domain;
using OfficeOpenXml;

namespace ConvertData.Infrastructure;

/// <summary>
/// Координирует чтение книги Excel, основной таблицы и объединение дополнительных листов.
/// </summary>
internal sealed class EpplusWorkbookReader
{
    private readonly EpplusWorksheetReader _worksheetReader = new();
    private readonly EpplusAdditionalSheetMerger _additionalSheetMerger = new();

    /// <summary>
    /// Считывает строки из книги Excel.
    /// </summary>
    /// <param name="path">Путь к файлу книги.</param>
    /// <returns>Список считанных строк.</returns>
    public List<Row> Read(string path)
    {
        using var package = new ExcelPackage(new FileInfo(path));
        var worksheet = package.Workbook.Worksheets
            .FirstOrDefault(x => string.Equals((x.Name ?? "").Trim(), "data", StringComparison.OrdinalIgnoreCase))
            ?? package.Workbook.Worksheets.FirstOrDefault();
        if (worksheet == null || worksheet.Dimension == null)
            return [];

        var result = _worksheetReader.Read(worksheet);
        if (result.IsMainTable)
            _additionalSheetMerger.Merge(package, worksheet, result.Rows);

        return result.Rows;
    }
}
