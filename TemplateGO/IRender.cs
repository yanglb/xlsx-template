using System.Text.Json;

namespace TemplateGO
{
    internal interface IRender
    {
        /// <summary>
        /// 渲染模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">数据</param>
        /// <param name="targetType">目标文件类型</param>
        public void Render(string templatePath, JsonElement data, string? targetType);
    }
}
