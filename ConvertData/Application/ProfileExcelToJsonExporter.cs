using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeOpenXml;

namespace ConvertData.Application;

internal sealed class ProfileExcelToJsonExporter
{
    /// <summary>
    /// Маппинг: заголовок Excel (case-insensitive) → JSON-поле(я).
    /// "Iy=Iz" → заполняет и Iy, и Iz одним значением.
    /// </summary>
    private static readonly Dictionary<string, string[]> HeaderMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["H"]     = ["H"],
        ["h"]     = ["H"],
        ["r1"]    = ["r1"],
        ["r2"]    = ["r2"],
        ["B"]     = ["B"],
        ["b"]     = ["B"],
        ["s"]     = ["s"],
        ["t"]     = ["t"],
        ["A"]     = ["A"],
        ["P"]     = ["P"],
        ["Iz"]    = ["Iz"],
        ["Iy"]    = ["Iy"],
        ["Ix"]    = ["Ix"],
        ["Iv"]    = ["Iv"],
        ["Iy=Iz"] = ["Iy", "Iz"],
        ["Iyz"]   = ["Iyz"],
        ["Wz"]    = ["Wz"],
        ["Wy"]    = ["Wy"],
        ["Wx"]    = ["Wx"],
        ["Wvo"]   = ["Wvo"],
        ["Sz"]    = ["Sz"],
        ["Sy"]    = ["Sy"],
        ["iz"]    = ["iz"],
        ["iy"]    = ["iy"],
        ["xo"]    = ["xo"],
        ["yo"]    = ["yo"],
        ["iu"]    = ["iu"],
        ["iv"]    = ["iv"],
    };

    /// <summary>
    /// Упорядоченный список всех геометрических полей в JSON.
    /// </summary>
    private static readonly string[] AllFields =
    [
        "H", "B", "s", "t", "r1", "r2",
        "A", "P",
        "Iz", "Iy", "Ix", "Iv", "Iyz",
        "Wz", "Wy", "Wx", "Wvo",
        "Sz", "Sy",
        "iz", "iy",
        "xo", "yo",
        "iu", "iv"
    ];

    /// <summary>
    /// Возможные заголовки столбца «Профиль» в Excel.
    /// </summary>
    private static readonly HashSet<string> ProfileHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Profile", "Профиль", "Сечение", "Наименование"
    };

    public void Export(string excelProfileDir, string outputJsonPath)
    {
        var files = new[] { "ProfileI.xlsx", "ProfileC.xlsx", "ProfileL.xlsx" };
        var allProfiles = new List<Dictionary<string, object>>();

        foreach (var fileName in files)
        {
            var filePath = Path.Combine(excelProfileDir, fileName);
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"  Пропущен (не найден): {fileName}");
                continue;
            }

            var profiles = ReadExcelFile(filePath);
            allProfiles.AddRange(profiles);
            Console.WriteLine($"  {fileName}: прочитано {profiles.Count} профилей");
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        Directory.CreateDirectory(Path.GetDirectoryName(outputJsonPath) ?? ".");
        File.WriteAllText(outputJsonPath, JsonSerializer.Serialize(allProfiles, options), Encoding.UTF8);
        Console.WriteLine($"  Итого записано {allProfiles.Count} профилей → {outputJsonPath}");
    }

    private static List<Dictionary<string, object>> ReadExcelFile(string path)
    {
        using var package = new ExcelPackage(new FileInfo(path));

        // Словарь профилей по имени (для объединения данных со всех листов)
        var profileMap = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        // Порядок профилей (для сохранения исходного порядка)
        var profileOrder = new List<string>();

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
                var profileName = (ws.Cells[r, startCol + profileCol].Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(profileName))
                    continue;

                if (!profileMap.TryGetValue(profileName, out var entry))
                {
                    entry = new Dictionary<string, object>
                    {
                        ["CONNECTION_GUID"] = Guid.NewGuid().ToString("D"),
                        ["Profile"] = profileName
                    };

                    foreach (var field in AllFields)
                        entry[field] = 0.0;

                    profileMap[profileName] = entry;
                    profileOrder.Add(profileName);
                }

                foreach (var (colIndex, jsonFields) in colMapping)
                {
                    var val = ReadDouble(ws, r, startCol + colIndex);
                    foreach (var field in jsonFields)
                        entry[field] = val;
                }
            }
        }

        return profileOrder.Select(name => profileMap[name]).ToList();
    }

    private static double ReadDouble(ExcelWorksheet ws, int row, int col)
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

        return 0.0;
    }
}