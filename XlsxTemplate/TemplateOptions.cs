using XlsxTemplate.Parser;

namespace XlsxTemplate
{
    /// <summary>
    /// 预加载图片
    /// </summary>
    /// <param name="image">图片内容，路径、url地址或base64内容</param>
    /// <param name="property">模板中的属性名</param>
    /// <param name="options">模板中的选项</param>
    /// <returns>处理后的图片内容</returns>
    public delegate string? PreLoadImageDelegate(string? image, string property, ImageOptions options);

    /// <summary>
    /// 数据转换
    /// </summary>
    /// <param name="value">输入值</param>
    /// <param name="options">选项</param>
    /// <returns>输出值</returns>
    public delegate object? TransformDelegate(object? value, TransformOptions options);

    /// <summary>
    /// 模板渲染选项
    /// </summary>
    public struct TemplateOptions
    {
        /// <summary>
        /// 预加载图片
        /// </summary>
        public PreLoadImageDelegate? PreLoadImage { get; set; }

        /// <summary>
        /// 转换器
        /// </summary>
        public Dictionary<string, TransformDelegate>? Transforms { get; set; }
    }
}
