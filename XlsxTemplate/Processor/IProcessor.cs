using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using XlsxTemplate.Parser;

namespace XlsxTemplate.Processor
{
    internal interface IProcessor
    {
        /// <summary>
        /// 处理模板内容
        /// </summary>
        /// <param name="processParams">参数</param>
        internal void Process(ProcessParams processParams);
    }

    /// <summary>
    /// 处理程序参数
    /// </summary>
    internal struct ProcessParams
    {
        public Cell Cell { get; set; }
        public string OriginValue { get; set; }
        public Grammar Parser { get; set; }
        public Sheet Sheet { get; set; }
        public JsonElement Data { get; set; }
        public SharedStringTable? SharedStringTable { get; set; }
        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }

        /// <summary>
        /// 是否在处理表格Table
        /// </summary>
        public bool InTable {  get; set; }

        public TemplateOptions Options { get; set; }
    }
}
