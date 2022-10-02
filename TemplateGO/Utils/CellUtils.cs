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
        /// 获取列引用 如 AB10 返回 AB
        /// </summary>
        /// <exception cref="ArgumentException"></exception>
        public static string ColumnReference(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) throw new ArgumentException("cellReference 为空");
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);
            return columnReference;
        }

        /// <summary>
        /// 获取列值，从1开始 如 A10 返回1, B5 返回 2
        /// </summary>
        public static int ColumnValue(string cellReference)
        {
            var columnReference = ColumnReference(cellReference);
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

        /// <summary>
        /// 判断两个区域是否相交
        /// </summary>
        /// <param name="range1">区域1 如: A1:C10</param>
        /// <param name="range2">区域2 如: B3:J20</param>
        /// <returns>是否相交</returns>
        public static bool IsIntersect(string range1, string range2)
        {
            return IsRowIntersect(range1, range2) && IsColumnIntersect(range1, range2);
        }

        /// <summary>
        /// 判断两个区域的行是否相交
        /// </summary>
        /// <param name="range1">区域1 如: A1:C10 或 1:10</param>
        /// <param name="range2">区域2 如: B3:J20 或 3:20</param>
        /// <returns>是否相交</returns>
        public static bool IsRowIntersect(string range1, string range2)
        {
            if (!Regex.IsMatch(range1, @"\w*\d+:\w*\d+"))
            {
                throw new ArgumentException($"“{range1}”不是有效的区域");
            }
            if (!Regex.IsMatch(range2, @"\w*\d+:\w*\d+"))
            {
                throw new ArgumentException($"“{range2}”不是有效的区域");
            }

            // 行
            var a = range1.Split(':');
            var b = range2.Split(':');
            var aMin = RowValue(a[0]);
            var aMax = RowValue(a[1]);
            var bMin = RowValue(b[0]);
            var bMax = RowValue(b[1]);

            return IsIntersect(aMin, aMax, bMin, bMax);
        }

        /// <summary>
        /// 判断两个区域的列是否相交
        /// </summary>
        /// <param name="range1">区域1 如: A1:C10 或 A:C</param>
        /// <param name="range2">区域2 如: B3:J20 或 A:C</param>
        /// <returns>是否相交</returns>
        public static bool IsColumnIntersect(string range1, string range2)
        {
            if (!Regex.IsMatch(range1, @"\w+\d*:\w+\d*"))
            {
                throw new ArgumentException($"“{range1}”不是有效的区域");
            }
            if (!Regex.IsMatch(range2, @"\w+\d*:\w+\d*"))
            {
                throw new ArgumentException($"“{range2}”不是有效的区域");
            }

            // 行
            var a = range1.Split(':');
            var b = range2.Split(':');
            var aMin = ColumnValue(a[0]);
            var aMax = ColumnValue(a[1]);
            var bMin = ColumnValue(b[0]);
            var bMax = ColumnValue(b[1]);

            return IsIntersect(aMin, aMax, bMin, bMax);
        }

        private static bool IsIntersect(int aMin, int aMax, int bMin, int bMax)
        {
            return ((aMin >= bMin && aMin <= bMax) || (aMax >= bMin && aMax <= bMax)) ||
                ((bMin >= aMin && bMin <= aMax) || (bMax >= aMin && bMax <= aMax));
        }

        /// <summary>
        /// 更新引用区域大小（起始值不变）
        /// </summary>
        /// <param name="origin">原始引用区域</param>
        /// <param name="shift">位移行数 + 表示增加 - 表示减少</param>
        /// <returns></returns>
        public static string UpdateToReference(string origin, int shift)
        {
            if (!Regex.IsMatch(origin, @"\w+\d+:\w+\d+")) throw new ArgumentException($"“{origin}”不是有效的区域");

            var from = origin.Split(':')[0];
            var to = origin.Split(':')[1];

            var toValue = RowValue(to) + shift;
            if (toValue <= 0) throw new ArgumentOutOfRangeException("超出有效范围");
            return $"{from}:{ColumnReference(to)}{toValue}";
        }
    }
}
