using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateGO.Parser;

namespace TemplateGO.Processor
{
    internal class ProcHyperlink : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 数据中指定的值
            object? value = GetValueByProperty(p.Data, p.Parser.Property);

            // 超链接内容 如果用引号 "" 包裹则当作字符串否则尝试从data中获取
            var options = new HyperlinkOptions(p.Parser.Options);
            object? content = options.Content;
            if(!string.IsNullOrEmpty(options.Content))
            {
                if (options.Content.StartsWith('"') && options.Content.EndsWith('"'))
                {
                    content = options.Content[1..^1];
                }
                else
                {
                    content = GetValueByProperty(p.Data, options.Content);
                }
            }

            // 设置单元格内容（空值）
            SetCellValue(p.Cell, p.OriginValue, p.Parser, content, p.SharedStringTable);

            // 无链接地址时清空
            if (value != null && value.GetType() == typeof(string) && !string.IsNullOrEmpty(value as string))
            {
                AddLink(p.Cell, p.WorksheetPart, $"{value}");
            }
            else
            {
                RemoveLink(p.Cell, p.WorksheetPart);
            }
        }

        private static void RemoveLink(Cell cell, WorksheetPart worksheetPart)
        {
            var hyperLinks = worksheetPart.Worksheet.Descendants<Hyperlinks>().FirstOrDefault();
            var link = hyperLinks?.Descendants<Hyperlink>().Where(item => item.Reference == cell.CellReference).FirstOrDefault();
            if (link == null) return;

            // 删除关系
            var hRelation = worksheetPart.HyperlinkRelationships.Where(item => item.Id == link.Id)?.FirstOrDefault();
            if (hRelation != null) worksheetPart.DeleteReferenceRelationship(hRelation);

            // 移除链接
            if (hyperLinks?.Count() <= 1)
            {
                hyperLinks.Remove();
            }
            link.Remove();
        }

        private static void AddLink(Cell cell, WorksheetPart worksheetPart, string linkUrl)
        {
            // 添加超链接
            var hyperLinks = worksheetPart.Worksheet.Descendants<Hyperlinks>().FirstOrDefault();
            if (hyperLinks == null)
            {
                hyperLinks = new Hyperlinks();

                // Worksheet 中需要将链接插入到 PageMargins 之前
                var pm = worksheetPart.Worksheet.Descendants<PageMargins>().FirstOrDefault();
                if (pm == null) worksheetPart.Worksheet.Append(hyperLinks);
                else worksheetPart.Worksheet.InsertBefore(hyperLinks, pm);
            }

            // 同一单元格只需要1个链接
            var link = hyperLinks.Descendants<Hyperlink>().Where(item => item.Reference == cell.CellReference).FirstOrDefault();
            if (link == null)
            {
                var id = $"Link{Guid.NewGuid().ToString().Split('-')[0]}";
                link = new Hyperlink()
                {
                    Reference = cell.CellReference,
                    Id = id,
                };
                hyperLinks.Append(link);
            }
            else
            {
                // 删除关系
                var hRelation = worksheetPart.HyperlinkRelationships.Where(item => item.Id == link.Id)?.FirstOrDefault();
                if (hRelation != null) worksheetPart.DeleteReferenceRelationship(hRelation);
            }

            // 重新设置关系
            worksheetPart.AddHyperlinkRelationship(new Uri(linkUrl, UriKind.Absolute), true, link.Id!);
        }
    }
}
