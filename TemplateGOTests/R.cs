using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace TemplateGOTests
{
    internal static class R
    {
        private static string AppDir
        {
            get
            {
                var file = Process.GetCurrentProcess().MainModule?.FileName;

                // 不存在返回当前路径
                if (string.IsNullOrEmpty(file)) return "";

                return Path.GetDirectoryName(file) ?? "";
            }
        }

        /// <summary>
        /// 获取相对于工程目录的文件（仅用于测试）
        /// </summary>
        internal static string FullPath(string relativeProjPath)
        {
            var path = Path.Join(AppDir, "../../..", relativeProjPath);
            return Path.GetFullPath(path);
        }

        /// <summary>
        /// 从测试数据中获取json内容
        /// </summary>
        /// <param name="file">JSON文件路径</param>
        /// <returns>JSON内容 RootElement</returns>
        internal static JsonElement JsonFromFile(string file)
        {
            var fullPath = R.FullPath(file);
            var jsonString = File.ReadAllText(fullPath);
            return JsonDocument.Parse(jsonString)!.RootElement;
        }
    }
}
