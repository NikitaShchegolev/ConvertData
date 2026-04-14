using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string path = @"..\ConvertData\EXCEL_Anchor\Anchor.xlsx";
        if (!File.Exists(path))
        {
            Console.WriteLine("File not found: " + path);
            return;
        }
        using var package = new ExcelPackage(new FileInfo(path));
        Console.WriteLine($"Sheets: {package.Workbook.Worksheets.Count}");
        foreach (var ws in package.Workbook.Worksheets)
        {
            Console.WriteLine($"\nWorksheet: {ws.Name}");
            if (ws.Dimension == null)
            {
                Console.WriteLine("  No data");
                continue;
            }
            int startRow = ws.Dimension.Start.Row;
            int endRow = Math.Min(startRow + 10, ws.Dimension.End.Row); // первые 10 строк
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;
            Console.WriteLine($"  Dimensions: {endRow - startRow + 1} rows, {endCol - startCol + 1} cols");
            // заголовки
            Console.Write("  Headers: ");
            for (int c = startCol; c <= Math.Min(startCol + 15, endCol); c++)
            {
                var val = ws.Cells[startRow, c].Text?.Trim();
                Console.Write($"[{c}] '{val}' ");
            }
            Console.WriteLine();
            // несколько строк данных
            for (int r = startRow + 1; r <= Math.Min(startRow + 5, endRow); r++)
            {
                Console.Write($"  Row {r}: ");
                for (int c = startCol; c <= Math.Min(startCol + 5, endCol); c++)
                {
                    var val = ws.Cells[r, c].Text?.Trim();
                    if (!string.IsNullOrEmpty(val))
                        Console.Write($"{val} ");
                }
                Console.WriteLine();
            }
        }
    }
}
