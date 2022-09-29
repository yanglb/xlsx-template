using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace TemplateGO.Utils
{
    public static class ImageUtils
    {
        /// <summary>
        /// 获取图片文件
        /// </summary>
        /// <param name="image">文件、网络地址、Base64编码的内容</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static string ToLocalFile(string image)
        {
            if (string.IsNullOrEmpty(image)) throw new ArgumentNullException("图片地址或内容不能为空");

            // Base64 编码
            // data:[mediatype][;base64],{base64-data}
            if (image.StartsWith("data:image")) return SaveBase64Image(image);

            // 网络图片
            if (Regex.IsMatch(image, @"^http(s)?://", RegexOptions.IgnoreCase))
            {
                return DownloadImageFile(image);
            }

            // 其它当作文件处理
            return image;
        }

        // 代码来自 ImagePartTypeInfo
        public static ImagePartType GetImagePartType(string extension)
        {
            switch (extension.ToLower())
            {
                case ".bmp":
                    return ImagePartType.Bmp;
                case ".emf":
                    return ImagePartType.Emf;
                case ".ico":
                    return ImagePartType.Icon;
                case ".jpg":
                    return ImagePartType.Jpeg;
                case ".jpeg":
                    return ImagePartType.Jpeg;
                case ".pcx":
                    return ImagePartType.Pcx;
                case ".png":
                    return ImagePartType.Png;
                case ".svg":
                    return ImagePartType.Svg;
                case ".tiff":
                    return ImagePartType.Tiff;
                case ".wmf":
                    return ImagePartType.Wmf;
                default:
                    throw new NotSupportedException(extension + " is not supported");
            }
        }

        /// <summary>
        /// 将Base64编码的内容保存到文件中
        /// </summary>
        /// <param name="imageBase64Data">data:[mediatype][;base64],{base64-data}</param>
        /// <returns>文件路径</returns>
        /// <exception cref="ArgumentException">不支持的类型时</exception>
        private static string SaveBase64Image(string imageBase64Data)
        {
            var idx = imageBase64Data.IndexOf(',');
            var meta = imageBase64Data[..idx];
            var base64 = imageBase64Data[(idx + 1)..];

            var imageType = Regex.Match(meta, @"data:(image/[\w-]+)").Groups[1].Value;
            if (string.IsNullOrEmpty(imageType)) throw new ArgumentException("不是有效的Base64编码图片。");
            var filePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + GetImageExtension(imageType));

            using (var fs = File.OpenWrite(filePath))
            {
                var data = Convert.FromBase64String(base64);
                fs.Write(data);
            }
            return filePath;
        }

        private static string DownloadImageFile(string imageUrl)
        {
            var task = DownloadImageFileAsync(imageUrl);
            task.Wait();
            return task.Result;
        }

        private static async Task<string> DownloadImageFileAsync(string imageUrl)
        {
            var client = new HttpClient();
            var res = await client.GetAsync(imageUrl);
            res.EnsureSuccessStatusCode();

            // 最终文件路径
            string filePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + "-" + Path.GetFileName(imageUrl));

            using var fs = File.Open(filePath, FileMode.Create);
            var st = await res.Content.ReadAsByteArrayAsync();
            fs.Write(st);
            return filePath;
        }

        private static string GetImageExtension(string imageType)
        {
            switch (imageType)
            {
                case "image/bmp":
                    return ".bmp";
                case "image/gif":
                    return ".gif";
                case "image/png":
                    return ".png";
                case "image/tiff":
                    return ".tiff";
                case "image/x-icon":
                    return ".icon";
                case "image/x-pcx":
                    return ".pcx";
                case "image/jpeg":
                    return ".jpeg";
                case "image/x-emf":
                    return ".emf";
                case "image/x-wmf":
                    return ".wmf";
                case "image/svg+xml":
                    return ".svg";
                default:
                    throw new ArgumentOutOfRangeException("imageType");
            }
        }

    }
}
