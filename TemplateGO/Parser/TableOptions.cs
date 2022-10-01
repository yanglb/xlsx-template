namespace TemplateGO.Parser
{
    public class TableOptions
    {
        public TableOptions() { }
        public TableOptions(Dictionary<string, string> options)
        {
            if (options.ContainsKey("titleCount"))
            {
                TitleCount = int.Parse(options["titleCount"]);
            }
            if(options.ContainsKey("keepExists"))
            {
                KeepExists = true;
            }
        }

        /// <summary>
        /// 标题行数目 key = titleCount, default = 1
        /// </summary>
        public long TitleCount { get; private set; } = 1;

        /// <summary>
        /// 是否清除原有数据 key = keepExists, default = false
        /// </summary>
        public bool KeepExists { get; private set; } = false;
    }
}
