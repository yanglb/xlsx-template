using XlsxTemplate.Utils;

namespace XlsxTemplate.Parser
{
    public class ImageOptions
    {
        public ImageOptions() { }
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
            if (options.ContainsKey("deleteMarked"))
            {
                DeleteMarked = true;
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

        /// <summary>
        /// 删除同单元格上的标记图片 key = deleteMarked
        /// 仅在添加了新图片后执行
        /// </summary>
        public bool DeleteMarked { get; private set; } = false;
    }
}
