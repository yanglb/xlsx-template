using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Text.RegularExpressions;
using TemplateGO.Parser;
using TemplateGO.Processor;
using TemplateGO.Transform;
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

        private bool processedTable = false;

        public void Render(string templatePath, JsonElement data, string? targetType, TemplateOptions options)
        {
            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(templatePath, true);

            // 后缀不一致时修改文档类型
            var newType = (Path.GetExtension(targetType)?.ToLower()) ?? "";
            if (string.Compare(Path.GetExtension(templatePath), newType, true) != 0)
            {
                Console.WriteLine($"转换文件类型为 {newType}");
                if (!SheetTypeMap.TryGetValue(newType, out SpreadsheetDocumentType value)) throw new ArgumentException($"未知的目标文档类型 {newType}");
                spreadsheetDocument.ChangeDocumentType(value);
                spreadsheetDocument.Save();
            }

            var workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null) throw new ArgumentException("模板格式错误: WorkbookPart 为空");
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;

            // 处理全部Sheet
            var sheets = workbookPart.Workbook.Descendants<Sheet>() ?? throw new ArgumentException("模板格式错误: Sheet 为空");
            foreach (var sheet in sheets)
            {
                ReplaceSheet(workbookPart, sheet, data, stringTable, options);
            }

            // 如果处理过表格则删除计算链
            if (processedTable)
            {
                var chainPart = workbookPart.CalculationChainPart;
                if (chainPart != null) workbookPart.DeletePart(chainPart);
            }

            // 处理 Text
            ReplaceText(workbookPart, data, options);

            // 保存工作表
            workbookPart.Workbook.Save();
            spreadsheetDocument.Save();
        }

        /// <summary>
        /// 替换除Cell外的数据
        /// </summary>
        private void ReplaceText(WorkbookPart workbookPart, JsonElement data, TemplateOptions options)
        {
            var texts = AllFlagedText(workbookPart);
            foreach (var text in texts)
            {
                var matchs = Grammar.Matches(text.Text);
                if (matchs == null || matchs.Count == 0)
                {
                    throw new Exception($"无法处理标记: {text.Text}");
                }

                foreach (var match in matchs.Cast<Match>())
                {
                    var parser = new Grammar(match.Value);
                    if (parser.Processor != ProcessorType.Value)
                    {
                        Console.WriteLine($"暂不支持处理单元格外的 {parser.Processor}");
                        continue;
                    }

                    var value = "";
                    try
                    {
                        var v = JsonUtils.GetValue(data, parser.Property);
                        v = ValueTransform.Transform(v, parser, options);
                        if (value.GetType() == typeof(JsonElement)) value = "[object Object]";
                        else value = $"{v}"; // 转为字符串
                    }
                    catch { }

                    // 只替换第1个满足条件的
                    var regex = new Regex(Regex.Escape(match.Value));
                    var newValue = regex.Replace(text.Text, value, 1);
                    text.Text = newValue;
                }
            }
        }

        private List<DocumentFormat.OpenXml.Drawing.Text> AllFlagedText(WorkbookPart workbookPart)
        {
            List<DocumentFormat.OpenXml.Drawing.Text> textList = new();

            // 来自图表的
            var charts = CellUtils.GetAllCharts(workbookPart);
            foreach (var chart in charts)
            {
                var texts = chart.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Where(r => Grammar.IsMatch(r.Text));
                textList.AddRange(texts);
            }

            // 来自Sheet内
            if (workbookPart.WorksheetParts != null)
            {
                foreach (var part in workbookPart.WorksheetParts)
                {
                    var texts = part.DrawingsPart?.WorksheetDrawing
                        .Descendants<DocumentFormat.OpenXml.Drawing.Text>()
                        .Where(r => Grammar.IsMatch(r.Text));
                    if (texts != null) textList.AddRange(texts);
                }
            }
            return textList;
        }

        private void ReplaceSheet(WorkbookPart workbookPart, Sheet sheet, JsonElement data, SharedStringTable? sharedStringTable, TemplateOptions options)
        {
            if (workbookPart.GetPartById(sheet.Id?.Value ?? "") is not WorksheetPart wsPart)
            {
                Console.WriteLine($"未发现该Sheet {sheet.Name} 的WorksheetPart内容。");
                return;
            }

            // 找出全部有效标识的单元格
            var cells = wsPart.Worksheet.Descendants<Cell>().Where(cell =>
            {
                var value = CellUtils.GetCellString(cell, sharedStringTable);
                return Grammar.IsMatch(value);
            });
            Console.WriteLine($"Sheet {sheet.Name} 中共发现 {cells?.Count() ?? 0} 个单元格需要处理。");
            if (cells == null) return;

            foreach (var cell in cells)
            {
                var cellValue = CellUtils.GetCellString(cell, sharedStringTable);
                var matchs = Grammar.Matches(cellValue);
                if (matchs == null)
                {
                    throw new Exception($"无法处理单元格: {cell.CellReference} => {cellValue}");
                }

                foreach (var match in matchs.Cast<Match>())
                {
                    var parser = new Grammar(match.Value);
                    var processor = Processor.Processor.ProcessorByTypeWithCache(parser.Processor);
                    if (processor == null) continue;
                    if (parser.Processor == ProcessorType.Table) processedTable = true;

                    // 处理
                    processor.Process(new ProcessParams()
                    {
                        Cell = cell,
                        OriginValue = cellValue,
                        Parser = parser,
                        Sheet = sheet,
                        Data = data,
                        SharedStringTable = sharedStringTable,
                        WorkbookPart = workbookPart,
                        WorksheetPart = wsPart,
                        Options = options,
                    });
                }
            }
        }
    }
}
