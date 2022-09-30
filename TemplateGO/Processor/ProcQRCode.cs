using dotnetCampus.OpenXmlUnitConverter;
using TemplateGO.Parser;
using TemplateGO.Utils;

namespace TemplateGO.Processor
{
    internal class ProcQRCode : ProcImage
    {
        protected new string GetImageLocalFile(string image, ImageOptions options)
        {
            return ImageUtils.CreateQrCode(image);
        }
    }
}
