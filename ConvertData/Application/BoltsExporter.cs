using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeOpenXml;

namespace ConvertData.Application;

internal sealed class BoltsExporter
{
    /// <summary>
    /// Маппинг: заголовок Excel (case-insensitive) → JSON-поле(я).
    /// Заголовки Excel могут быть без суффикса _Anchor, с пробелами в конце.
    /// </summary>
    private static readonly Dictionary<string, string[]> HeaderMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["GOST_bolts"] = ["GOST_bolts"],
        ["d_SP43"] = ["d_SP43"],
        ["Ab_SP43"] = ["Ab_SP43"],
        ["Abn_SP43"] = ["Abn_SP43"]
    };

    /// <summary>
    /// Упорядоченный список всех числовых полей в JSON (для инициализации нулями).
    /// </summary>
    private static readonly string[] NumericFields =
    [
        "d_SP43",
        "Ab_SP43",
        "Abn_SP43"
    ];

    /// <summary>
    /// Возможные заголовки столбца «Профиль» в Excel (для анкеров это Connect_Name_Anchor или CONNECTION_CODE_Anchor).
    /// </summary>
    private static readonly HashSet<string> ProfileHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "d"
    };

    private static readonly Dictionary<string, string> FileCategoryMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["TableBoltsSP43.xlsx"] = "Анкерные болты",
    };

    public void Export(string excelAnchorDir, string outputJsonPath)
    {
        var files = new[] { "TableBoltsSP43.xlsx" };
        var allAnchors = new List<Dictionary<string, object>>();

        foreach (var fileName in files)
        {
            var filePath = Path.Combine(excelAnchorDir, fileName);
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"  Пропущен (не найден): {fileName}");
                continue;
            }

            var category = FileCategoryMap.GetValueOrDefault(fileName, "");
            var anchors = ReadExcelFile(filePath, category);
            allAnchors.AddRange(anchors);
            Console.WriteLine($"  {fileName}: прочитано {anchors.Count} болтов");
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, JsonSerializer.Serialize(allAnchors, options), Encoding.UTF8);
        Console.WriteLine($"  Итого записано {allAnchors.Count} болтов → {outputJsonPath}");
    }

    private static List<Dictionary<string, object>> ReadExcelFile(string path, string category)
    {
        using var package = new ExcelPackage(new FileInfo(path));

        // Словарь анкеров по имени (для объединения данных со всех листов)
        var anchorMap = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        // Порядок анкеров (для сохранения исходного порядка)
        var anchorOrder = new List<string>();

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
                var anchorName = (ws.Cells[r, startCol + profileCol].Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(anchorName))
                    continue;

                if (!anchorMap.TryGetValue(anchorName, out var entry))
                {
                    // Создаём запись с вложенным объектом Geometry
                    var geometry = new Dictionary<string, object>();
                    foreach (var field in NumericFields)
                        geometry[field] = 0.0;

                    // Создаём запись с корневыми полями в нужном порядке
                    var rootFields = new Dictionary<string, object>
                    {
                        ["CONNECTION_GUID"] = Guid.NewGuid().ToString("D"),
                        ["Category"] = category
                    };
                    entry = new Dictionary<string, object>();
                    // Добавляем корневые поля в нужном порядке
                    foreach (var kv in rootFields)
                        entry[kv.Key] = kv.Value;
                    // Добавляем Geometry после корневых полей
                    entry["Geometry"] = geometry;

                    // Если столбец профиля соответствует полю из HeaderMap, добавляем его значение
                    var profileHeader = headers[profileCol];
                    if (HeaderMap.TryGetValue(profileHeader, out var profileFields))
                    {
                        var profileVal = ReadCellValue(ws, r, startCol + profileCol);
                        foreach (var field in profileFields)
                        {
                            if (NumericFields.Contains(field))
                                geometry[field] = profileVal;
                            else
                                entry[field] = profileVal;
                        }
                    }

                    anchorMap[anchorName] = entry;
                    anchorOrder.Add(anchorName);
                }

                // Получаем ссылку на geometry
                var geometryDict = (Dictionary<string, object>)entry["Geometry"];

                foreach (var (colIndex, jsonFields) in colMapping)
                {
                    var val = ReadCellValue(ws, r, startCol + colIndex);
                    foreach (var field in jsonFields)
                    {
                        // Определяем, куда поместить поле: в корень или в Geometry
                        if (NumericFields.Contains(field))
                            geometryDict[field] = val;
                        else
                            entry[field] = val;
                    }
                }
            }
        }

        // Преобразуем каждую запись в итоговый формат (убедимся, что Geometry присутствует)
        var result = new List<Dictionary<string, object>>();
        foreach (var name in anchorOrder)
        {
            var entry = anchorMap[name];
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