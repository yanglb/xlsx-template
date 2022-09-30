using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace TemplateGO.Utils
{
    public static class CellUtils
    {
        /// <summary>
        /// 获取单元格字符内容
        /// </summary>
        public static string GetCellString(Cell cell, SharedStringTable? sharedStringTable)
        {
            var value = cell.InnerText;
            if (string.IsNullOrEmpty(value)) return "";

            // 从 SharedStringTable 中获取
            if (cell.DataType != null && cell.DataType == CellValues.SharedString && sharedStringTable != null)
            {
                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }

        /// <summary>
        /// 获取行值，从1开始 如 A10 返回10, B5 返回 5
        /// </summary>
        public static int RowValue(string cellReference)
        {
            var match = Regex.Match(cellReference, @"(\d+)");
            if (!match.Success) throw new ArgumentException("cellReference 为空");
            return int.Parse(match.Groups[1].Value);
        }

        /// <summary>
        /// 获取列值，从1开始 如 A10 返回1, B5 返回 2
        /// </summary>
        public static int ColumnValue(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) throw new ArgumentException("cellReference 为空");
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);
                mulitplier *= 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }

        /// <summary>
        /// 获取单元格宽度 Emu 单位
        /// TODO: 后期处理
        /// </summary>
        public static long CellWidth(Cell cell, WorksheetPart worksheetPart)
        {
            var colValue = ColumnValue(cell.CellReference!);
            var col = worksheetPart.Worksheet.Descendants<Column>()?.Where(
                r => r.Min?.Value <= colValue && r.Max?.Value >= colValue
                && r.CustomWidth?.Value == true
                ).FirstOrDefault();

            // 未设置列宽时返回默认值
            if (col == null)
            {
                return 0;
            }
            return 1;
        }
    }
}
