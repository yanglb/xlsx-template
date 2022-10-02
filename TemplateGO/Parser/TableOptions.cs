namespace TemplateGO.Parser
{
    public class TableOptions
    {
        public TableOptions() { }
        public TableOptions(Dictionary<string, string> options)
        {
            if (options.ContainsKey("titleCount"))
            {
                TitleCount = uint.Parse(options["titleCount"]);
            }
            if(options.ContainsKey("sampleCount"))
            {
                SampleCount = uint.Parse(options["sampleCount"]);
            }

            if (options.ContainsKey("limit"))
            {
                Limit = int.Parse(options["limit"]);
            }
        }

        /// <summary>
        /// 标题行数目 key = titleCount, default = 1
        /// </summary>
        public uint TitleCount { get; private set; } = 1;

        /// <summary>
        /// 样本数据量 key = sampleCount, default = 0
        /// 添加完数据后会删除样本数据行
        /// </summary>
        public uint SampleCount { get; private set; } = 0;

        /// <summary>
        /// 数量限制 key = limit
        /// 为空时不限制
        /// </summary>
        public int? Limit { get; private set; } = null;
    }
}
