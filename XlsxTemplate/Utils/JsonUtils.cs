using System.Text.Json;
using System.Text.RegularExpressions;

namespace XlsxTemplate.Utils
{
    public class JsonUtils
    {
        /// <summary>
        /// 根据属性在 JsonElement 中获取数据
        /// paths 为空时返回 data
        /// 
        /// <code>
        /// var data = { 'a': [{ 'b': { 'c': 3 } }, 4], 'd': 'hello' }
        /// GetValue(data, "a[0].b.c") // => 3
        /// GetValue(data, "d") // => "hello"
        /// </code>
        /// </summary>
        /// <param name="data">JSON数据</param>
        /// <param name="paths">属性路径名</param>
        /// <returns></returns>
        public static object? GetValue(JsonElement data, string? paths)
        {
            if(string.IsNullOrEmpty(paths)) return data;

            var keys = new Queue<string>(paths.Split('.'));
            return DoGetValue(data, keys);
        }

        /// <summary>
        /// 获取JSON值
        /// 
        /// ValueKind 为 Number 时有小数返回 double 否则返回 int32
        /// Array/Object 直接返回JsonElement
        /// </summary>
        public static object? GetValue(JsonElement value)
        {
            // 获取最终结果
            switch (value.ValueKind)
            {
                case JsonValueKind.Number:
                    // 有小数返回 Double 否则返回 int
                    if (value.GetRawText().Contains('.')) return value.GetDouble();
                    return value.GetInt32();

                case JsonValueKind.True:
                case JsonValueKind.False:
                    return value.GetBoolean();

                case JsonValueKind.String:
                    return value.GetString()!;

                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
            }

            // Array/Object 直接返回
            return value;
        }

        static object? DoGetValue(JsonElement data, Queue<string> keys)
        {
            var key = keys.Dequeue();
            if (key == null) throw new ArgumentException("Key 不能为空");

            JsonElement value;
            if (Regex.IsMatch(key, @"\[\d+\]$"))
            {
                var match = Regex.Match(key, @"(\w+)?\[(\d+)\]")!;
                if (!match.Success || match.Groups.Count != 3)
                {
                    throw new ArgumentException($"key 格式不正确 {key}");
                }

                // 分别获取 key 及 数组索引
                key = match.Groups[1]?.Value;
                var index = int.Parse(match.Groups[2]?.Value!);

                if (!string.IsNullOrEmpty(key)) value = data.GetProperty(key);
                else value = data;

                if (value.ValueKind != JsonValueKind.Array) throw new ArgumentException($"{key} 指定的对象不是数组。");
                value = value.EnumerateArray().ElementAt(new Index(index));
            }
            else
            {
                value = data.GetProperty(key);
            }

            // 后续还有key
            if (keys.Count > 0) return DoGetValue(value, keys);

            // 获取最终结果
            return GetValue(value);
        }
    }
}
