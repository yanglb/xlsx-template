using TemplateGO.Processor;

namespace TemplateGO
{
    public class Parser
    {
        /// <summary>
        /// 解析器语法 ${key[|proc[:[settingKey1=settingValue1],[settingKey2=settingValue2]]}
        /// </summary>
        public Parser(string content)
        {
            if (!content.StartsWith("${") || !content.EndsWith("}"))
            {
                throw new ArgumentException($"语法错误: {content}");
            }

            // 删除 ${} 字符
            var values = content[2..^1];

            Origin = content;

            var prop = values.Split('|');

            // 属性key
            Property = prop[0] ?? "";

            // 处理器
            var proc = prop.Length > 1 ? prop[1]! : ProcessorType.Value;
            var opts = proc.Split(':');
            Processor = opts[0]!;

            // 选项
            Options = new Dictionary<string, string>();
            if (opts.Length > 1)
            {
                foreach (var item in opts[1].Split(','))
                {
                    if (item == null) continue;
                    var kvs = item.Split('=');
                    Options.Add(kvs[0], kvs.Length > 1 ? kvs[1] : "");
                }
            }
        }

        /// <summary>
        /// 原始值
        /// </summary>
        public string Origin { get; private set; }

        /// <summary>
        /// 属性名
        /// </summary>
        public string Property { get; private set; }

        /// <summary>
        /// 处理器
        /// </summary>
        public string Processor { get; private set; }

        public Dictionary<string, string> Options { get; private set; }
    }
}
