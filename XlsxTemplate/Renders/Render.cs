namespace XlsxTemplate.Renders
{
    internal static class Render
    {
        private static Dictionary<string, Type> RenderType = new Dictionary<string, Type>()
        {
            {".xlsx", typeof(Spreadsheet) },
            {".xltx", typeof(Spreadsheet) },
            {".xlsm", typeof(Spreadsheet) },
            {".xltm", typeof(Spreadsheet) },
            {".xlam", typeof(Spreadsheet) },
        };

        /// <summary>
        /// 通过文件扩展名获取对应的渲染实例
        /// </summary>
        /// <param name="extension">扩展名，包括“.”符号。</param>
        /// <returns></returns>
        public static IRender RenderByExtension(string extension) {
        var type = RenderType[extension];
            if (type == null) throw new ArgumentException($"未定义 {extension} 类型的渲染处理程序。");

            var ins = Activator.CreateInstance(type) as IRender;
            if (ins == null) throw new Exception($"无法实例化用于处理 {extension} 模板的渲染程序。");
            return ins!;
        }
    }
}
