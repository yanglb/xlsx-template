namespace TemplateGO.Parser
{
    public class HyperlinkOptions
    {
        public HyperlinkOptions() { }
        public HyperlinkOptions(Dictionary<string, string> options)
        {
            if (options.ContainsKey("content"))
            {
                Content = options["content"];
            }
        }

        /// <summary>
        /// 超连接文字属性 key = content
        /// </summary>
        public string? Content { get; private set; }
    }
}
