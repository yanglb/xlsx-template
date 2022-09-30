using dotnetCampus.OpenXmlUnitConverter;
using System.Text.RegularExpressions;

namespace TemplateGO.Utils
{
    public static class ParserUtils
    {
        /// <summary>
        /// 根据选项解析值和单位 如 1.23cm 值为1.23 单位为 cm
        /// </summary>
        /// <param name="input">输入内容</param>
        /// <param name="value">值输出</param>
        /// <param name="unit">单位输出</param>
        /// <exception cref="ArgumentException"></exception>
        public static void ValueWithUnit(string input, out string value, out string unit)
        {
            var match = Regex.Match(input.Trim(), @"^(-?[\d.]+)([\w\W]*)$");
            if (!match.Success) throw new ArgumentException($"无法解析 {input}");
            value = match.Groups[1].Value;
            unit = match.Groups[2].Value;
        }

        /// <summary>
        /// 将单位转换为 Emu
        /// 支持 cm/px/in
        /// </summary>
        /// <param name="value">输入值</param>
        /// <param name="unit">原单位</param>
        /// <returns>Emu单位的数值</returns>
        /// <exception cref="ArgumentException"></exception>
        public static long ValueToEmu(string value, string unit)
        {
            switch (unit)
            {
                case "cm":
                    {
                        var convert = new Cm(double.Parse(value));
                        return (long)convert.ToEmu().Value;
                    }

                case "px":
                    {
                        var convert = new Pixel(double.Parse(value));
                        return (long)convert.ToEmu().Value;
                    }

                case "in":
                    {
                        var convert = new Inch(double.Parse(value));
                        return (long)convert.ToEmu().Value;
                    }

                default:
                    throw new ArgumentException($"无法将 {unit} 转换为Emu");
            }
        }

        /// <summary>
        /// 解析输入值到 Emu 单位
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static long ParseValueToEmu(string input)
        {
            ValueWithUnit(input, out string value, out string unit);
            return ValueToEmu(value, unit);
        }
    }
}
