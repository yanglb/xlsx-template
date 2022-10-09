using TemplateGO.Parser;

namespace TemplateGO
{
    public struct TransformOptions
    {
        /// <summary>
        /// 属性名
        /// </summary>
        public string Property { get; set; }

        /// <summary>
        /// 处理程序，参见 ProcessorType
        /// </summary>
        public string Processor { get; set; }

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
