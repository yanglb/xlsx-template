using System.Text.Json;

namespace TemplateGO
{
    public static class TemplateGO
    {
        public static void Render(string templatePath, JsonElement data, string saveTo)
        {
            // 检查模板类型
            var srcExtension = Path.GetExtension(templatePath).ToLower();
            var trgExtension = Path.GetExtension(saveTo).ToLower();
            if (string.IsNullOrEmpty(srcExtension) || string.IsNullOrEmpty(trgExtension)) {
                throw new ArgumentException("未知模板类型。");
            }

            // 获取渲染实例
            var render = Renders.Render.RenderByExtension(srcExtension);

            // 处理模板文件
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + "-" + Path.GetFileName(templatePath));
            File.Copy(templatePath, tempFile);

            // 渲染
            render.Render(tempFile, data, trgExtension);

            // 直接移动到目标目录即可
            File.Move(tempFile, saveTo, true);
        }
    }
}
