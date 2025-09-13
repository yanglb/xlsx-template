using XlsxTemplate.Parser;

namespace XlsxTemplate
{
    public struct TransformOptions
    {
        /// <summary>
        /// 当前执行的转换器
        /// </summary>
        public string Transform { get; set; }

        /// <summary>
        /// 语法/标记块内容
        /// </summary>
        public Grammar Grammar { get; set; }
    }
}
