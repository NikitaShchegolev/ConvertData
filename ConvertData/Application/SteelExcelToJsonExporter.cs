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
    /// <summary>
    /// Маппинг: заголовок Excel (case-insensitive) → JSON-поле(я).
    /// Заголовки Excel могут быть без суффикса _Anchor, с пробелами в конце.
    /// </summary>
    private static readonly Dictionary<string, string[]> HeaderMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Steel"] = ["Steel_Mark"],
        ["Name_Table"] = ["Name_Table"],
        ["tmin"] = ["tmin"],
        ["tmax"] = ["tmax"],
        ["Ryn"] = ["Ryn"],
        ["Run"] = ["Run"],
        ["Ry"] = ["Ry"],
        ["Ru"] = ["Ru"],
        ["if"] = ["if"],
    };

    // /// <summary>
    // /// Упорядоченный список всех числовых полей в JSON (для инициализации нулями).
    // /// </summary>
    // private static readonly string[] NumericFields =
    // [
    //     "Steel_Mark",
    //     "Name_Table",
    //     "tmin",
    //     "tmax",
    //     "Ryn",
    //     "Run",
    //     "Ry",
    //     "Ru",
    //     "if"
    // ];

    /// <summary>
    /// Возможные заголовки столбца «Профиль» в Excel (для анкеров это Connect_Name_Anchor или CONNECTION_CODE_Anchor).
    /// </summary>
    private static readonly HashSet<string> ProfileHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Steel_Mark", "Steel Mark"
    };

    private static readonly Dictionary<string, string> FileCategoryMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["MarkSteel.xlsx"] = "Марка стали",
    };

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

            var category = FileCategoryMap.GetValueOrDefault(fileName, "");
            var steel = ReadExcelFile(filePath, category);
            allSteels.AddRange(steel);
            Console.WriteLine($"  {fileName}: прочитано {steel.Count} анкеров");
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, JsonSerializer.Serialize(allSteels, options), Encoding.UTF8);
        Console.WriteLine($"  Итого записано {allSteels.Count} анкеров → {outputJsonPath}");
    }

    private static List<Dictionary<string, object>> ReadExcelFile(string path, string category)
    {
        using var package = new ExcelPackage(new FileInfo(path));

        // Словарь анкеров по имени (для объединения данных со всех листов)
        var steelMap = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        // Порядок анкеров (для сохранения исходного порядка)
        var steelOrder = new List<string>();

        foreach (var ws in package.Workbook.Worksheets)
        {
            if (ws.Dimension == null)
                continue;

            int startRow = ws.Dimension.Start.Row;
            int endRow = ws.Dimension.End.Row;
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;

            // --- заголовки ---
            var headers = new string[endCol - startCol + 1];
            for (int c = startCol; c <= endCol; c++)
                headers[c - startCol] = (ws.Cells[startRow, c].Text ?? "").Trim();

            // --- столбец «Профиль» ---
            int profileCol = -1;
            for (int i = 0; i < headers.Length; i++)
            {
                if (ProfileHeaders.Contains(headers[i]))
                {
                    profileCol = i;
                    break;
                }
            }
            if (profileCol < 0)
                profileCol = 0;

            // --- маппинг столбцов → JSON-полей ---
            var colMapping = new List<(int colIndex, string[] jsonFields)>();
            for (int i = 0; i < headers.Length; i++)
            {
                if (i == profileCol || string.IsNullOrWhiteSpace(headers[i]))
                    continue;
                if (HeaderMap.TryGetValue(headers[i], out var fields))
                    colMapping.Add((i, fields));
            }

            if (colMapping.Count == 0)
                continue;

            // --- данные ---
            for (int r = startRow + 1; r <= endRow; r++)
            {
                var steelName = (ws.Cells[r, startCol + profileCol].Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(steelName))
                    continue;

                if (!steelMap.TryGetValue(steelName, out var entry))
                {
                    // Создаём вложенные объекты согласно новой структуре
                    var steel = new Dictionary<string, object>();
                    var geometry = new Dictionary<string, object>();
                    var tension = new Dictionary<string, object>();
                    var variable = new Dictionary<string, object>();

                    // Инициализируем все числовые поля нулями
                    steel["Steel_Mark"] = "";
                    steel["Name_Table"] = "";
                    geometry["tmin"] = 0.0;
                    geometry["tmax"] = 0.0;
                    tension["Ryn"] = 0.0;
                    tension["Run"] = 0.0;
                    tension["Ry"] = 0.0;
                    tension["Ru"] = 0.0;
                    variable["if"] = 0.0;

                    // Создаём запись с корневыми полями
                    entry = new Dictionary<string, object>
                    {
                        ["CONNECTION_GUID"] = Guid.NewGuid().ToString("D"),
                        ["Steel"] = steel,
                        ["Geometry"] = geometry,
                        ["Tension"] = tension,
                        ["Variable"] = variable
                    };

                    // Если столбец профиля соответствует полю из HeaderMap, добавляем его значение
                    var profileHeader = headers[profileCol];
                    if (HeaderMap.TryGetValue(profileHeader, out var profileFields))
                    {
                        var profileVal = ReadCellValue(ws, r, startCol + profileCol);
                        foreach (var field in profileFields)
                        {
                            // Определяем, в какой объект поместить поле
                            if (field == "Steel_Mark" || field == "Name_Table")
                                steel[field] = profileVal;
                            else if (field == "tmin" || field == "tmax")
                                geometry[field] = profileVal;
                            else if (field == "Ryn" || field == "Run" || field == "Ry" || field == "Ru")
                                tension[field] = profileVal;
                            else if (field == "if")
                                variable[field] = profileVal;
                            else
                                entry[field] = profileVal;
                        }
                    }

                    steelMap[steelName] = entry;
                    steelOrder.Add(steelName);
                }

                // Получаем ссылки на вложенные объекты
                var steelDict = (Dictionary<string, object>)entry["Steel"];
                var geometryDict = (Dictionary<string, object>)entry["Geometry"];
                var tensionDict = (Dictionary<string, object>)entry["Tension"];
                var variableDict = (Dictionary<string, object>)entry["Variable"];

                foreach (var (colIndex, jsonFields) in colMapping)
                {
                    var val = ReadCellValue(ws, r, startCol + colIndex);
                    foreach (var field in jsonFields)
                    {
                        // Определяем, в какой объект поместить поле
                        if (field == "Steel_Mark" || field == "Name_Table")
                            steelDict[field] = val;
                        else if (field == "tmin" || field == "tmax")
                            geometryDict[field] = val;
                        else if (field == "Ryn" || field == "Run" || field == "Ry" || field == "Ru")
                            tensionDict[field] = val;
                        else if (field == "if")
                            variableDict[field] = val;
                        else
                            entry[field] = val;
                    }
                }
            }
        }

        // Преобразуем каждую запись в итоговый формат (убедимся, что Geometry присутствует)
        var result = new List<Dictionary<string, object>>();
        foreach (var name in steelOrder)
        {
            var entry = steelMap[name];
            // Если какие-то корневые поля отсутствуют, они уже добавлены
            result.Add(entry);
        }
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