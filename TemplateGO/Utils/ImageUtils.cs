using DocumentFormat.OpenXml.Packaging;
using dotnetCampus.OpenXmlUnitConverter;
using SixLabors.ImageSharp;
using System.Text.RegularExpressions;
using TemplateGO.Parser;

namespace TemplateGO.Utils
{
    /// <summary>
    /// 以 EMU 为单位
    /// </summary>
    public struct ImageShapeInfo
    {
        public long X;
        public long Y;
        public long W;
        public long H;
    }

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
        /// 获取合适的图片位置及大小
        /// </summary>
        /// <param name="options">选项</param>
        /// <param name="imageFile">图片文件</param>
        public static ImageShapeInfo GetImageShape(ImageOptions options, string imageFile)
        {
            using var imgInfo = Image.Load(imageFile);
            return GetImageShape(options, imgInfo.Width, imgInfo.Height);
        }

        /// <summary>
        /// 获取合适的图片位置及大小
        /// </summary>
        /// <param name="options">选项</param>
        /// <param name="width">图片宽度（px）</param>
        /// <param name="height">图片高度（px）</param>
        public static ImageShapeInfo GetImageShape(ImageOptions options, int width, int height)
        {
            var wEmu = (long)new Pixel(width).ToEmu().Value;
            var hEmu = (long)new Pixel(height).ToEmu().Value;

            var imageShape = new ImageShapeInfo();
            imageShape.X = options.Padding;
            imageShape.Y = options.Padding;

            // 未指定宽度及高度 使用图片大小
            if (options.FrameWidth == null && options.FrameHeight == null)
            {
                imageShape.W = wEmu;
                imageShape.H = hEmu;
                return imageShape;
            }

            long? w = null, h = null;
            if (options.FrameWidth != null && options.FrameHeight != null)
            {
                // 同时设置长宽时长边完全显示、短边缩放
                var rateImg = width * 1.0 / height;
                var rateFrm = options.FrameWidth!.Value * 1.0 / options.FrameHeight!.Value;

                // 图片比外框要宽
                if (rateImg > rateFrm) w = options.FrameWidth - (options.Padding * 2);
                // 图片比外框高
                else h = options.FrameHeight!.Value - (options.Padding * 2);
            }
            else if (options.FrameWidth != null)
            {
                // 仅指定宽度
                w = options.FrameWidth - (options.Padding * 2);
            }
            else if (options.FrameHeight != null)
            {
                // 仅指定高度
                h = options.FrameHeight - (options.Padding * 2);
            }

            // 以宽度为准
            if (w != null)
            {
                imageShape.W = w.Value;
                if (imageShape.W <= 0) throw new Exception("设置边距后图片宽度不足");
                var rate = imageShape.W * 1.0 / width;
                imageShape.H = (long)Math.Round(rate * height, MidpointRounding.AwayFromZero);
                return imageShape;
            }

            // 以高度为准
            if (h != null)
            {
                imageShape.H = h.Value;
                if (imageShape.H <= 0) throw new Exception("设置边距后图片高度不足");
                var rate = imageShape.H * 1.0 / height;
                imageShape.W = (long)Math.Round(rate * width, MidpointRounding.AwayFromZero);
                return imageShape;
            }

            return imageShape;
        }

        /// <summary>
        /// 移动到中心点
        /// </summary>
        /// <param name="shape">形状</param>
        /// <param name="fw">外框宽度</param>
        /// <param name="fh">外框高度</param>
        public static ImageShapeInfo MoveToCenter(ImageShapeInfo shape, long fw, long fh)
        {
            shape.X = (long)(fw / 2.0 - shape.W / 2.0);
            shape.Y = (long)(fh / 2.0 - shape.H / 2.0);
            if (shape.X < 0) shape.X = 0;
            if (shape.Y < 0) shape.Y = 0;
            return shape;
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
