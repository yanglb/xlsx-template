using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateGO.Processor
{
    internal class ProcImage : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 数据中指定的值
            object? value = GetValueByProperty(p.Data, p.Parser.Property);

            // 设置单元格内容（空值）
            SetCellValue(p.Cell, p.OriginValue, p.Parser, null, p.SharedStringTable);

            // 无数据时清空
            if (value != null && value.GetType() == typeof(string) && !string.IsNullOrEmpty(value as string))
            {
                AddImage(p.Cell, p.WorksheetPart, $"{value}");
            }
            else
            {
                RemoveImage(p.Cell, p.WorksheetPart);
            }
        }

        private void RemoveImage(Cell cell, WorksheetPart worksheetPart) { }

        private void AddImage(Cell cell, WorksheetPart worksheetPart, string image) { }
    }
}
