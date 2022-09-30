using TemplateGO.Utils;

namespace TemplateGO.Parser
{
    public class ImageOptions
    {
        public ImageOptions(Dictionary<string, string> options)
        {
            if (options.ContainsKey("padding"))
            {
                Padding = ParserUtils.ParseValueToEmu(options["padding"]);
            }
            if(options.ContainsKey("fw"))
            {
                FrameWidth = ParserUtils.ParseValueToEmu(options["fw"]);
            }
            if (options.ContainsKey("fh"))
            {
                FrameHeight = ParserUtils.ParseValueToEmu(options["fh"]);
            }
        }

        /// <summary>
        /// 内边距 key = fw, default = 0
        /// </summary>
        public long Padding { get; private set; }

        /// <summary>
        /// 外框宽度 key = fw
        /// </summary>
        public long? FrameWidth { get; private set; }

        /// <summary>
        /// 外框高度 key = fh
        /// </summary>
        public long? FrameHeight { get; private set; }
    }
}
