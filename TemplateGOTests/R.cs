using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using TemplateGO.Utils;

namespace TemplateGOTests
{
    internal static class R
    {
        private static string AppDir
        {
            get
            {
                var executablePath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
                return executablePath ?? "";
            }
        }

        /// <summary>
        /// 获取相对于工程目录的文件（仅用于测试）
        /// </summary>
        internal static string FullPath(string relativeProjPath)
        {
            var path = Path.Combine(AppDir, "../../..", relativeProjPath);
            return Path.GetFullPath(path);
        }

        /// <summary>
        /// 从测试数据中获取json内容
        /// </summary>
        /// <param name="file">JSON文件路径</param>
        /// <returns>JSON内容 RootElement</returns>
        internal static JsonElement JsonFromFile(string file)
        {
            var fullPath = R.FullPath(file);
            var jsonString = File.ReadAllText(fullPath);
            return JsonDocument.Parse(jsonString)!.RootElement;
        }

        /// <summary>
        /// 获取输出文件地址
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        internal static string OutFullPath(string file)
        {
            return FullPath(Path.Combine("out", file));
        }

        internal static string? CellStringValue(SpreadsheetDocument doc, string sheetName, string cellReference)
        {
            var sheet = doc.WorkbookPart?.Workbook.Descendants<Sheet>().FirstOrDefault((r) => r.Name == sheetName);
            if (sheet == null) throw new ArgumentException($"工作表中不存在 {sheetName}。");

            var sheetPart = doc.WorkbookPart?.GetPartById(sheet.Id!) as WorksheetPart;
            if (sheetPart == null) throw new ArgumentException("不存在 WorksheetPart");

            var shareStringTable = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

            var cell = sheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(r => r.CellReference == cellReference);
            if (cell == null) throw new ArgumentException($"{sheetName} 中不存在 {cellReference}");

            return CellUtils.GetCellString(cell, shareStringTable);
        }

        internal static string[] ColumnStrings(Worksheet worksheet, SharedStringTable? sharedStringTable, string columnName, int startRow, int? endRow = null)
        {
            var res = worksheet.Descendants<Cell>()
                .Where(r =>
                    r.CellReference!.Value!.StartsWith(columnName) &&
                    (r.Parent as Row)!.RowIndex!.Value >= startRow &&
                    (endRow == null || (r.Parent as Row)!.RowIndex!.Value <= endRow)
                )
                .Select(r => CellUtils.GetCellString(r, sharedStringTable))
                .ToArray();

            return res;
        }
    }
}
