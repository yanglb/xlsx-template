using System.Text.RegularExpressions;
using TemplateGO.Processor;

namespace TemplateGO.Parser
{
    /// <summary>
    /// 语法块
    /// <code>${key[|proc[|transform1|transform2][:[settingKey1=settingValue1],[settingKey2=settingValue2]]}</code>
    /// </summary>
    public class Grammar
    {
        /// <summary>
        /// 是否为语法标记块
        /// </summary>
        /// <param name="input">输入内容</param>
        /// <returns>是否为标记块</returns>
        public static bool IsMatch(string? input)
        {
            if (string.IsNullOrEmpty(input)) return false;
            return Regex.IsMatch(input, @"\${[^}]*}+");
        }

        /// <summary>
        /// 获取所有符合的标记块
        /// </summary>
        /// <param name="input">输入内容</param>
        /// <returns>符合要求的标记块</returns>
        public static MatchCollection? Matches(string? input)
        {
            if (string.IsNullOrEmpty(input)) return null;

            var matchs = Regex.Matches(input, @"\${([^}]*)*}");
            if (matchs.Count == 0) return null;
            return matchs;
        }

        /// <summary>
        /// 解析器语法
        /// </summary>
        public Grammar(string content)
        {
            if (!content.StartsWith("${") || !content.EndsWith("}"))
            {
                throw new ArgumentException($"语法错误: {content}");
            }

            // 原始内容
            Origin = content;

            // 删除 ${} 字符
            var values = content[2..^1];

            // 属性|处理器|转换器 : 选项
            var optsIdx = values.IndexOf(':');

            // 属性、处理器、转换器
            var prop = (optsIdx >= 0 ? values[..optsIdx] : values).Split('|');

            // 属性key
            Property = (prop[0] ?? "").Trim();

            // 处理器
            Processor = prop.Length > 1 ? prop[1]!.Trim() : "";
            if (string.IsNullOrEmpty(Processor)) Processor = ProcessorType.Value;

            // 剩余均为转换器
            Transforms = prop.Length > 2 ? prop[2..].Select(r => r.Trim()).ToArray() : Array.Empty<string>();

            // 选项
            var opts = optsIdx >= 0 ? values[(optsIdx + 1)..]?.Trim() : "";
            Options = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(opts))
            {
                foreach (var item in opts.Split(','))
                {
                    if (item == null) continue;
                    var kvs = item.Split('=');
                    Options.Add(kvs[0].Trim(), kvs.Length > 1 ? kvs[1].Trim() : "");
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
        /// 转换器
        /// </summary>
        public string[] Transforms { get; private set; }

        /// <summary>
        /// 处理器
        /// </summary>
        public string Processor { get; private set; }

        public Dictionary<string, string> Options { get; private set; }
    }
}
