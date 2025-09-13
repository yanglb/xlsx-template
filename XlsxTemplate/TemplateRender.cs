﻿using System.Text.Json;

namespace XlsxTemplate
{
    public static class TemplateRender
    {
        /// <summary>
        /// 渲染模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="jsonString">JSON字符数据</param>
        /// <param name="saveTo">输出文件</param>
        public static void Render(string templatePath, string jsonString, string saveTo)
        {
            Render(templatePath, jsonString, saveTo, new TemplateOptions());
        }

        /// <summary>
        /// 渲染模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="jsonString">JSON字符数据</param>
        /// <param name="saveTo">输出文件</param>
        /// <param name="options">选项</param>
        public static void Render(string templatePath, string jsonString, string saveTo, TemplateOptions options)
        {
            var jd = JsonDocument.Parse(jsonString)!.RootElement;
            Render(templatePath, jd, saveTo, options);
        }

        /// <summary>
        /// 渲染模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">数据</param>
        /// <param name="saveTo">输出文件</param>
        public static void Render(string templatePath, JsonElement data, string saveTo)
        {
            Render(templatePath, data, saveTo, new TemplateOptions());
        }

        /// <summary>
        /// 渲染模板
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">数据</param>
        /// <param name="saveTo">输出文件</param>
        /// <param name="options">选项</param>
        public static void Render(string templatePath, JsonElement data, string saveTo, TemplateOptions options)
        {
            // 检查模板类型
            var srcExtension = Path.GetExtension(templatePath).ToLower();
            var trgExtension = Path.GetExtension(saveTo).ToLower();
            if (string.IsNullOrEmpty(srcExtension) || string.IsNullOrEmpty(trgExtension))
            {
                throw new ArgumentException("未知模板类型。");
            }

            // 获取渲染实例
            var render = Renders.Render.RenderByExtension(srcExtension);

            // 处理模板文件
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + "-" + Path.GetFileName(templatePath));
            File.Copy(templatePath, tempFile);

            // 渲染
            render.Render(tempFile, data, trgExtension, options);

            // 直接移动到目标目录即可
            var path = Path.GetDirectoryName(saveTo);
            if(!string.IsNullOrEmpty(path) && !Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            File.Move(tempFile, saveTo, true);
        }
    }
}
