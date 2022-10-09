using TemplateGO.Parser;

namespace TemplateGO.Transform
{
    internal class ValueTransform
    {
        public static object? Transform(object? value, Grammar grammar, TemplateOptions options)
        {
            // 无需转换
            if (options.Transforms == null) return value;
            if (grammar.Transforms.Length == 0) return value;

            // 按顺序处理
            foreach (var transform in grammar.Transforms)
            {
                if (string.IsNullOrEmpty(transform) || !options.Transforms.ContainsKey(transform))
                {
                    Console.WriteLine($"不存在转换器: {transform}");
                    continue;
                }

                var trans = options.Transforms[transform];
                value = trans(value, new TransformOptions()
                {
                    Grammar = grammar,
                    Transform = transform,
                });
            }
            return value;
        }
    }
}
