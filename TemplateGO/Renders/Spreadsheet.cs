using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Text.RegularExpressions;
using TemplateGO.Processor;
using TemplateGO.Utils;

namespace TemplateGO.Renders
{
    /// <summary>
    /// 电子表格模板处理程序
    /// </summary>
    internal class Spreadsheet : IRender
    {
        /// <summary>
        /// 支持的文件种类
        /// </summary>
        private static readonly Dictionary<string, SpreadsheetDocumentType> SheetTypeMap = new()
        {
            { ".xlsx", SpreadsheetDocumentType.Workbook },
            { ".xltx", SpreadsheetDocumentType.Template },
            { ".xlsm", SpreadsheetDocumentType.Template },
            { ".xltm", SpreadsheetDocumentType.MacroEnabledTemplate },
            { ".xlam", SpreadsheetDocumentType.AddIn },
        };

        public void Render(string templatePath, JsonElement data, string? targetType)
        {
            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(templatePath, true);

            // 后缀不一致时修改文档类型
            var newType = (Path.GetExtension(targetType)?.ToLower()) ?? "";
            if (string.Compare(Path.GetExtension(templatePath), newType, true) != 0)
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

        /// <summary>
        /// 处理器缓存
        /// </summary>
        private Dictionary<string, IProcessor?> ProcessorCache = new();
        private IProcessor? ProcessorByType(string type)
        {
            if (!ProcessorCache.ContainsKey(type))
            {
                var processor = Processor.Processor.ProcessorByType(type);
                ProcessorCache[type] = processor;
            }

            return ProcessorCache[type];
        }

        private void ReplaceSheet(WorkbookPart workbookPart, Sheet sheet, JsonElement data, SharedStringTable? sharedStringTable)
        {
            if (workbookPart.GetPartById(sheet.Id?.Value ?? "") is not WorksheetPart wsPart)
            {
                Console.WriteLine($"未发现该Sheet {sheet.Name} 的WorksheetPart内容。");
                return;
            }

            // 找出全部有效标识的单元格
            // ${key[|proc[:[settingKey1=settingValue1],[settingKey2=settingValue2]]}
            var cells = wsPart.Worksheet.Descendants<Cell>().Where((Func<Cell, bool>)(cell =>
            {
                var value = CellUtils.GetCellString(cell, sharedStringTable);
                if (string.IsNullOrEmpty(value)) return false;
                return Regex.IsMatch(value, @"\${[^}]+}+");
            }));
            Console.WriteLine($"Sheet {sheet.Name} 中共发现 {cells?.Count() ?? 0} 个单元格需要处理。");
            if (cells == null) return;

            foreach (var cell in cells)
            {
                var cellValue = CellUtils.GetCellString(cell, sharedStringTable);
                var matchs = Regex.Matches(cellValue, @"\${([^}]+)+}");
                if (matchs == null || matchs.Count == 0)
                {
                    throw new Exception($"无法处理单元格: {cell.CellReference} => {cellValue}");
                }

                foreach (Match match in matchs)
                {
                    var parser = new Parser(match.Value);
                    var processor = ProcessorByType(parser.Processor);
                    if (processor == null) continue;

                    // 处理
                    // cell, cellValue, parser, sheet, data, sharedStringTable
                    processor.Process(new ProcessParams()
                    {
                        Cell = cell,
                        OriginValue = cellValue,
                        Parser = parser,
                        Sheet = sheet,
                        Data = data,
                        SharedStringTable = sharedStringTable,
                        WorkbookPart = workbookPart,
                        WorksheetPart = wsPart
                    });
                }
            }
        }
    }
}
