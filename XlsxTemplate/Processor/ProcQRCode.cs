using XlsxTemplate.Parser;
using XlsxTemplate.Utils;

namespace XlsxTemplate.Processor
{
    internal class ProcQRCode : ProcImage
    {
        protected override string? GetImageLocalFile(string? image, ImageOptions options, ProcessParams p)
        {
            if (string.IsNullOrEmpty(image)) return null;
            return ImageUtils.CreateQrCode(image);
        }
    }
}
