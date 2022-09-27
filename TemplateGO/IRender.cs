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
        public void Render(string templatePath, JsonElement data);
    }
}
