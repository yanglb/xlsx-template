using System.Diagnostics;
using System.IO;

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
    }
}
