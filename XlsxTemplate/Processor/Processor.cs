namespace XlsxTemplate.Processor
{
    internal class Processor
    {
        private static Dictionary<string, Type> ProcessorTypeMap = new Dictionary<string, Type>()
        {
            {ProcessorType.Value, typeof(ProcValue) },
            {ProcessorType.Hyperlink, typeof(ProcHyperlink) },
            {ProcessorType.Image, typeof(ProcImage) },
            {ProcessorType.QRCode, typeof(ProcQRCode) },
            {ProcessorType.Table, typeof(ProcTable) },
        };

        /// <summary>
        /// 获取处理程序
        /// </summary>
        /// <param name="procType">处理器种类</param>
        public static IProcessor? ProcessorByType(string procType)
        {
            if (!ProcessorTypeMap.TryGetValue(procType, out Type? type))
            {
                Console.WriteLine($"未定义 {procType} 处理程序。");
                return null;
            }

            if (Activator.CreateInstance(type) is not IProcessor ins)
            {
                throw new Exception($"无法实例化用于处理 {procType} 的处理程序。");
            }
            return ins;
        }

        // 处理器缓存
        private static readonly Dictionary<string, IProcessor?> ProcessorCache = [];

        /// <summary>
        /// 获取处理程序(支持缓存)
        /// </summary>
        /// <param name="type">处理器种类</param>
        public static IProcessor? ProcessorByTypeWithCache(string type)
        {
            if (!ProcessorCache.TryGetValue(type, out IProcessor? value))
            {
                var processor = ProcessorByType(type);
                value = processor;
                ProcessorCache[type] = value;
            }

            return value;
        }
    }
}
