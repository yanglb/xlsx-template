using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace TemplateGO.Renders
{
    /// <summary>
    /// 电子表格模板处理程序
    /// </summary>
    internal class Spreadsheet : IRender
    {
        /// <summary>
        /// Excel 种类
        /// </summary>
        private static Dictionary<string, SpreadsheetDocumentType> SheetTypeMap = new()
        {
            { ".xlsx", SpreadsheetDocumentType.Workbook },
            { ".xltx", SpreadsheetDocumentType.Template },
            { ".xlsm", SpreadsheetDocumentType.Template },
            { ".xltm", SpreadsheetDocumentType.MacroEnabledTemplate },
            { ".xlam", SpreadsheetDocumentType.AddIn },
        };

        public void Render(string templatePath, JsonElement data, string? targetType)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(templatePath, true))
            {
                // 后缀不一致时修改文档类型
                var newType = (Path.GetExtension(targetType)?.ToLower())??"";
                if (string.Compare( Path.GetExtension(templatePath), newType, true) != 0)
                {
                    Console.WriteLine($"转换文件类型为 {newType}");
                    if (!SheetTypeMap.ContainsKey(newType)) throw new ArgumentException($"未知的目标文档类型 {newType}");
                    spreadsheetDocument.ChangeDocumentType(SheetTypeMap[newType]);
                    spreadsheetDocument.Save();
                }

                var workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart == null) throw new ArgumentException("模板格式错误: WorkbookPart 为空");
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;

                // 处理全部Sheet
                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                if (sheets == null) throw new ArgumentException("模板格式错误: Sheet 为空");
                foreach (var sheet in sheets)
                {
                    ReplaceSheet(workbookPart, sheet, data, stringTable);
                }

                // 保存工作表
                workbookPart.Workbook.Save();
                spreadsheetDocument.Save();
            }
        }

        private static void ReplaceSheet(WorkbookPart workbookPart, Sheet sheet, JsonElement data, SharedStringTable? sharedStringTable)
        {
            WorksheetPart? wsPart = workbookPart.GetPartById(sheet.Id?.Value ?? "") as WorksheetPart;

            // 找出全部有效标识的单元格
            // ${key[|proc[:[settingKey1=settingValue1],[settingKey2=settingValue2]]}
            var cells = wsPart?.Worksheet.Descendants<Cell>().Where(cell =>
            {
                var value = CellString(cell, sharedStringTable);
                if (string.IsNullOrEmpty(value)) return false;
                return Regex.IsMatch(value, @"\${[^}]+}+");
            });
            Console.WriteLine($"Sheet {sheet.Name} 中共发现 {cells?.Count() ?? 0} 个单元格需要处理。");
            if (cells == null) return;

            foreach (var cell in cells)
            {
                var cellValue = CellString(cell, sharedStringTable);
                var matchs = Regex.Matches(cellValue, @"\${([^}]+)+}");
                if (matchs == null || matchs.Count == 0)
                {
                    throw new Exception($"无法处理单元格: {cell.CellReference} => {cellValue}");
                }
                // 暂时只处理第1个
                if (matchs.Count == 1)
                {
                    var match = matchs[0]!;
                    var prop = match.Groups[1].Value?.Split('|')!;

                    // 属性key
                    var key = prop[0]!;

                    // 处理器
                    var proc = prop.Count() > 1 ? prop[1]! : null;
                    object? value = null;
                    try
                    {
                        value = Utils.GetValue(data, key);
                    }
                    catch (Exception)
                    {
                    }

                    // 空值
                    if (value == null)
                    {
                        cell.CellValue = null;
                        cell.DataType = null;
                    }

                    else if (value.GetType() == typeof(int))
                    {
                        cell.CellValue = new CellValue((int)value);
                        cell.DataType = CellValues.Number;
                    }
                    else if (value.GetType() == typeof(double))
                    {
                        cell.CellValue = new CellValue((double)value);
                        cell.DataType = CellValues.Number;
                    }
                    else if (value.GetType() == typeof(bool))
                    {
                        cell.CellValue = new CellValue((bool)value);
                        cell.DataType = CellValues.Boolean;
                    }
                    else if (value.GetType() == typeof(string))
                    {
                        cell.CellValue = new CellValue((string)value);
                        cell.DataType = CellValues.String;
                    }

                    // 无法解析的对象
                    else/*if (value.GetType() == typeof(JsonElement))*/
                    {
                        cell.CellValue = new CellValue("[object Object]");
                        cell.DataType = CellValues.String;
                    }
                }
                //Console.WriteLine($"cell {cell.CellReference}: {cellValue}");

                //cell.CellValue = new CellValue(12.5);
                //cell.DataType = CellValues.Number;
            }
        }

        private static string CellString(Cell cell, SharedStringTable? sharedStringTable)
        {
            var value = cell.InnerText;
            if (string.IsNullOrEmpty(value)) return "";

            // 从 SharedStringTable 中获取
            if (cell.DataType != null && cell.DataType == CellValues.SharedString && sharedStringTable != null)
            {
                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
