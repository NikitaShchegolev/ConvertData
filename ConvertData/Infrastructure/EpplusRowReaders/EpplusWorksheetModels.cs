using ConvertData.Domain;

namespace ConvertData.Infrastructure;

/// <summary>
/// Результат чтения листа Excel с признаком основной таблицы.
/// </summary>
internal sealed record EpplusWorksheetReadResult(List<Row> Rows, bool IsMainTable);

/// <summary>
/// Границы заполненной области листа Excel.
/// </summary>
internal sealed record EpplusWorksheetBounds(int StartRow, int EndRow, int StartCol, int EndCol);
