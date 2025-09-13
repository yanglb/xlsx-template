using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Text.RegularExpressions;
using TemplateGO.Parser;
using TemplateGO.Transform;
using TemplateGO.Utils;

namespace TemplateGO.Processor
{
    internal class BaseProcess
    {
        protected void SetCellValue(Cell cell, string originValue, Grammar parser, object? value, SharedStringTable? sharedStringTable)
        {
            // 设置单元格内容
            if (originValue == parser.Origin)
            {
                // 仅在原始值与解析器中的原始值一致时调整格式，其它情况均视为字符
                SetCellValueWithType(cell, value);
            }
            else
            {
                // 使用字符替换方式
                SetCellValuePart(cell, parser.Origin, value, sharedStringTable);
            }
        }

        /// <summary>
        /// 获取内容并执行转换方法
        /// </summary>
        /// <param name="data">JSON数据</param>
        /// <param name="parser">标记语法解析结果</param>
        /// <param name="options">选项</param>
        protected object? GetValueAndTransform(JsonElement data, Grammar parser, TemplateOptions options)
        {
            object? value = GetValueByProperty(data, parser.Property);
            value = ValueTransform.Transform(value, parser, options);
            return value;
        }

        /// <summary>
        /// 获取无需转换的内容
        /// </summary>
        /// <param name="data">JSON数据</param>
        /// <param name="property">属性名，如果包裹在又引号中则视为直接内容</param>
        protected object? GetValueByProperty(JsonElement data, string? property)
        {
            try
            {
                // 如果属性包裹在""中则视为值
                if (!string.IsNullOrEmpty(property))
                {
                    var p = property.Trim();
                    if (p.StartsWith('"') && p.EndsWith('"')) return p[1..^1];
                }
                // 否则获取内容
                return JsonUtils.GetValue(data, property);
            }
            // 数据中不存在时视为 null
            catch (KeyNotFoundException) { }

            // 均返回null
            return null;
        }

        protected void SetCellValueWithType(Cell cell, object? value)
        {
            // 空值
            if (value == null)
            {
                cell.CellValue = null;
                cell.DataType = null;
            }

            else if (value.GetType() == typeof(int))
            {
                cell.CellValue = new CellValue((int)value);
                cell.DataType = CellValues.Number;
            }
            else if (value.GetType() == typeof(uint))
            {
                cell.CellValue = new CellValue(Convert.ToInt32(value));
                cell.DataType = CellValues.Number;
            }
            else if (value.GetType() == typeof(double))
            {
                cell.CellValue = new CellValue((double)value);
                cell.DataType = CellValues.Number;
            }
            else if (value.GetType() == typeof(bool))
            {
                cell.CellValue = new CellValue((bool)value);
                cell.DataType = CellValues.Boolean;
            }
            else if (value.GetType() == typeof(string))
            {
                cell.CellValue = new CellValue((string)value);
                cell.DataType = CellValues.String;
            }

            // 无法解析的对象
            else/*if (value.GetType() == typeof(JsonElement))*/
            {
                cell.CellValue = new CellValue("[object Object]");
                cell.DataType = CellValues.String;
            }
        }

        protected void SetCellValuePart(Cell cell, string match, object? value, SharedStringTable? sharedStringTable)
        {
            var currentValue = CellUtils.GetCellString(cell, sharedStringTable);
            var newValue = ReplaceCellValue(currentValue, match, value);

            // 设置值及类型
            cell.CellValue = new CellValue(newValue);
            cell.DataType = CellValues.String;
        }

        protected string ReplaceCellValue(string currentValue, string match, object? value)
        {
            if (value == null) value = "";
            if (value.GetType() == typeof(JsonElement)) value = "[object Object]";
            if (value.GetType() == typeof(bool)) value = (bool)value ? "TRUE" : "FALSE";

            // 只替换第1个满足条件的
            var regex = new Regex(Regex.Escape(match));
            return regex.Replace(currentValue, $"{value}", 1);
        }
    }
}
