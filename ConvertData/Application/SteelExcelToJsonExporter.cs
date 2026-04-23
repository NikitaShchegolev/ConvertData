using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

using OfficeOpenXml;

namespace ConvertData.Application;

internal sealed class SteelExcelToJsonExporter
{
    public void Export(string excelSteelDir, string outputJsonPath)
    {
        var files = new[] { "MarkSteel.xlsx" };
        var allSteels = new List<Dictionary<string, object>>();

        foreach (var fileName in files)
        {
            var filePath = Path.Combine(excelSteelDir, fileName);
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"  Пропущен (не найден): {fileName}");
                continue;
            }

            var steel = ReadExcelFile(filePath);
            allSteels.AddRange(steel);
            Console.WriteLine($"  {fileName}: прочитано {steel.Count} записей");
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, JsonSerializer.Serialize(allSteels, options), Encoding.UTF8);
        Console.WriteLine($"  Итого записано {allSteels.Count} записей → {outputJsonPath}");
    }

    private static List<Dictionary<string, object>> ReadExcelFile(string path)
    {
        using var package = new ExcelPackage(new FileInfo(path));
        var result = new List<Dictionary<string, object>>();

        foreach (var ws in package.Workbook.Worksheets)
        {
            if (ws.Dimension == null)
                continue;

            int startRow = ws.Dimension.Start.Row;
            int endRow = ws.Dimension.End.Row;
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;

            Console.WriteLine($"  Лист '{ws.Name}': строк {endRow - startRow}");

            // Читаем заголовки
            var headers = new string[endCol - startCol + 1];
            for (int c = startCol; c <= endCol; c++)
                headers[c - startCol] = (ws.Cells[startRow, c].Text ?? "").Trim();

            // Определяем индексы столбцов по заголовкам
            int steelCol = -1, nameTableCol = -1, tminCol = -1, tmaxCol = -1;
            int rynCol = -1, runCol = -1, ryCol = -1, ruCol = -1, caseParamCol = -1;

            //for (int i = 0; i < headers.Length; i++)
            //{
            //    if (ProfileHeaders.Contains(headers[i]))
            //    {
            //        profileCol = i;
            //        break;
            //    }
            //}
            //if (profileCol < 0)
            //    profileCol = 0;

            // Проверяем, что все необходимые столбцы найдены
            if (steelCol == -1)
            {
                Console.WriteLine("  ОШИБКА: не найден столбец 'Steel'");
                continue;
            }

            // Обрабатываем все строки данных
            for (int r = startRow + 1; r <= endRow; r++)
            {
                // Читаем значение стали
                var steelVal = (ws.Cells[r, startCol + steelCol].Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(steelVal))
                    continue; // Пропускаем полностью пустые строки

                // Создаём запись
                var entry = new Dictionary<string, object>
                {
                    ["CONNECTION_GUID"] = Guid.NewGuid().ToString("D")
                };

                // Создаём вложенные объекты
                var steel = new Dictionary<string, object>();
                var geometry = new Dictionary<string, object>();
                var tension = new Dictionary<string, object>();
                var variable = new Dictionary<string, object>();

                // Заполняем Steel_Mark
                steel["Steel_Mark"] = steelVal;

                // Заполняем остальные поля, если столбцы найдены
                if (nameTableCol != -1)
                    steel["Name_Table"] = ReadCellValue(ws, r, startCol + nameTableCol);
                else
                    steel["Name_Table"] = "";

                if (tminCol != -1)
                    geometry["tmin"] = ReadCellValue(ws, r, startCol + tminCol);
                else
                    geometry["tmin"] = 0.0;

                if (tmaxCol != -1)
                    geometry["tmax"] = ReadCellValue(ws, r, startCol + tmaxCol);
                else
                    geometry["tmax"] = 0.0;

                if (rynCol != -1)
                    tension["Ryn"] = ReadCellValue(ws, r, startCol + rynCol);
                else
                    tension["Ryn"] = 0.0;

                if (runCol != -1)
                    tension["Run"] = ReadCellValue(ws, r, startCol + runCol);
                else
                    tension["Run"] = 0.0;

                if (ryCol != -1)
                    tension["Ry"] = ReadCellValue(ws, r, startCol + ryCol);
                else
                    tension["Ry"] = 0.0;

                if (ruCol != -1)
                    tension["Ru"] = ReadCellValue(ws, r, startCol + ruCol);
                else
                    tension["Ru"] = 0.0;

                if (caseParamCol != -1)
                {
                    var caseVal = ReadCellValue(ws, r, startCol + caseParamCol);
                    // Преобразуем в строку, если нужно
                    variable["caseParametr_t"] = caseVal?.ToString() ?? "";
                }
                else
                {
                    variable["caseParametr_t"] = "";
                }

                // Добавляем вложенные объекты в запись
                entry["Steel"] = steel;
                entry["Geometry"] = geometry;
                entry["Tension"] = tension;
                entry["Variable"] = variable;

                result.Add(entry);
            }
        }

        Console.WriteLine($"  Итого обработано строк: {result.Count}");
        return result;
    }

    private static object ReadCellValue(ExcelWorksheet ws, int row, int col)
    {
        var raw = ws.Cells[row, col].Value;
        if (raw is double d) return d;
        if (raw is int i) return i;
        if (raw is decimal dec) return (double)dec;
        if (raw is float f) return f;

        var text = raw?.ToString()?.Trim() ?? "";
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;
        if (double.TryParse(text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            return result;
        // Если не число, возвращаем строку
        return text;
    }
}