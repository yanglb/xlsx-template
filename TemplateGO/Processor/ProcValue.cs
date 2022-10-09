namespace TemplateGO.Processor
{
    internal class ProcValue : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 数据中指定的值
            object? value = GetValueAndTransform(p.Data, p.Parser, p.Options);

            // 设置单元格内容
            SetCellValue(p.Cell, p.OriginValue, p.Parser, value, p.SharedStringTable);
        }
    }
}
