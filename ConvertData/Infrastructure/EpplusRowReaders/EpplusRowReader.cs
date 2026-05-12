using ConvertData.Application;
using ConvertData.Domain;
using ConvertData.Infrastructure.Interop;
using ConvertData.Infrastructure.Parsing;

namespace ConvertData.Infrastructure;

/// <summary>
/// Читает Excel-файлы через EPPlus и при необходимости конвертирует legacy .xls во временный .xlsx.
/// </summary>
internal sealed class EpplusRowReader : IRowReader
{
    /// <summary>
    /// Считывает строки из Excel-файла.
    /// </summary>
    /// <param name="path">Путь к входному Excel-файлу.</param>
    /// <returns>Список считанных строк.</returns>
    public List<Row> Read(string path)
    {
        var format = ExcelFileSignature.Detect(path);
        if (format == ExcelFileFormat.ZipXlsx)
            return new EpplusWorkbookReader().Read(path);

        var tmpXlsx = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(path) + "_converted_" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            ExcelXlsConverter.ConvertXlsToXlsxViaExcel(path, tmpXlsx);
            if (!File.Exists(tmpXlsx))
                throw new InvalidDataException("Failed to convert .xls to .xlsx (temporary file not created)");

            return new EpplusWorkbookReader().Read(tmpXlsx);
        }
        finally
        {
            try { if (File.Exists(tmpXlsx)) File.Delete(tmpXlsx); } catch { }
        }
    }
}
