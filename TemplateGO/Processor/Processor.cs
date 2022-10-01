namespace TemplateGO.Processor
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
            if (!ProcessorTypeMap.ContainsKey(procType))
            {
                Console.WriteLine($"未定义 {procType} 处理程序。");
                return null;
            }

            var type = ProcessorTypeMap[procType];
            if (Activator.CreateInstance(type) is not IProcessor ins)
            {
                throw new Exception($"无法实例化用于处理 {procType} 的处理程序。");
            }
            return ins;
        }
    }
}
