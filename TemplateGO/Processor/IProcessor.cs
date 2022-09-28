using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;

namespace TemplateGO.Processor
{
    internal interface IProcessor
    {
        /// <summary>
        /// 处理模板内容
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="originValue">单元格原始值</param>
        /// <param name="parser">解析器</param>
        /// <param name="sheet">所属Sheet</param>
        /// <param name="data">数据</param>
        internal void Process(Cell cell, string originValue, Parser parser, Sheet sheet, JsonElement data, SharedStringTable? sharedStringTable);
    }
}
