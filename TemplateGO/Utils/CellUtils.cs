using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;

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

        public static List<Chart> GetAllCharts(WorkbookPart workbookPart)
        {
            var chartList = new List<Chart>();

            // 图表Sheet
            foreach (var chartSheetPart in workbookPart.ChartsheetParts)
            {
                var cs = chartSheetPart.DrawingsPart?.ChartParts;
                if (cs == null) continue;
                foreach (var cc in cs)
                {
                    chartList.AddRange(cc.ChartSpace.Descendants<Chart>());
                }
            }

            // 普通Sheet中的图表
            foreach (var sheetPart in workbookPart.WorksheetParts)
            {
                var cs = sheetPart.DrawingsPart?.ChartParts;
                if (cs == null) continue;
                foreach (var cc in cs)
                {
                    chartList.AddRange(cc.ChartSpace.Descendants<Chart>());
                }
            }
            return chartList;
        }

        public static string ExpandRowReference(string reference, int rowShift)
        {
            var idx = reference.LastIndexOf(':');
            if (idx == -1) throw new ArgumentException($"{reference} 不是有效的公式引用");

            var pre = reference[..(idx + 1)]!;
            var row = reference[(idx + 1)..]!;

            var ends = Regex.Match(row, @"^(\$?[^\d\$]+\$?)(\d+)$");
            if (!ends.Success) throw new ArgumentException($"{reference} 不是有效的公式引用");

            return $"{pre}{ends.Groups[1].Value}{int.Parse(ends.Groups[2].Value) + rowShift}";
        }

        /// <summary>
        /// 更新图表引用（仅处理引用表格区域图表的行数变化）
        /// </summary>
        /// <param name="workbookPart">工作表对象</param>
        /// <param name="sheetName">数据变化的表格</param>
        /// <param name="rowShift">增加/减少的行数</param>
        public static void UpdateChartReference(WorkbookPart workbookPart, string sheetName, int rowShift, string originArea)
        {
            // 偏移为0时也需要更新
            //if (rowShift == 0) return;
            var chartList = GetAllCharts(workbookPart);

            foreach (var chart in chartList)
            {
                if (chart.PlotArea == null) continue;

                // 所有引用 sheetName 的公式
                var formulas = chart.Descendants<Formula>().Where((f) =>
                {
                    var formula = f.InnerText;
                    if (
                        string.IsNullOrEmpty(formula) || // 不能为空
                        !formula.Contains(sheetName) ||  // 必需引用该Sheet
                        !formula.Contains(':')           // 必需是区域类型的
                    )
                    {
                        return false;
                    }

                    // 位于 originArea 区域内
                    var area = formula[(formula.IndexOf('!') + 1)..];
                    area = area.Replace("$", string.Empty);
                    return IsIntersect(area, originArea);
                });
                foreach (var f in formulas)
                {
                    UpdateReference(f, rowShift);
                }
            }
        }

        /// <summary>
        /// 更新迷你图数据引用
        /// 引用 originArea 的更新引用区域
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <param name="shift"></param>
        /// <param name="range"></param>
        public static void UpdateSparklineReference(WorkbookPart workbookPart, string sheetName, int shift, string originArea)
        {
            // 所有引用 originArea 的迷你图
            var refCharts = new List<Sparkline>();
            foreach (var sheetPart in workbookPart.WorksheetParts)
            {
                var lines = sheetPart.Worksheet.Descendants<Sparkline>().Where((sLine) =>
                {
                    var formula = sLine.Formula?.InnerText;
                    if (
                        string.IsNullOrEmpty(formula) || // 不能为空
                        !formula.Contains(sheetName) ||  // 必需引用该Sheet
                        !formula.Contains(':')           // 必需是区域类型的
                    )
                    {
                        return false;
                    }

                    // 位于 originArea 区域内
                    var area = formula[(formula.IndexOf('!') + 1)..];
                    return IsIntersect(area, originArea);

                });
                refCharts.AddRange(lines);
            }

            foreach (var chart in refCharts)
            {
                var newFormula = new DocumentFormat.OpenXml.Office.Excel.Formula(ExpandRowReference(chart.Formula!.InnerText!, shift));
                chart.Formula = newFormula;
            }
        }

        /// <summary>
        /// 更新迷你图位置区域
        /// 当前工作表中位于 originArea 后迷你图
        /// </summary>
        /// <param name="worksheetPart"></param>
        /// <param name="shift"></param>
        /// <param name="originArea"></param>
        public static void UpdateSparklineArea(WorksheetPart worksheetPart, int shift, string originArea)
        {
            // 所有位于 originArea 以后的图表
            var maxRow = RowValue(originArea.Split(':')[1]);

            // 需要更新区域的图表
            var spartLines = worksheetPart.Worksheet.Descendants<Sparkline>().Where((line) =>
            {
                var celRef = line.ReferenceSequence?.InnerText;
                if (string.IsNullOrEmpty(celRef)) return false;
                return RowValue(celRef) > maxRow;
            });

            foreach (var spartLine in spartLines)
            {
                var col = ColumnReference(spartLine.ReferenceSequence?.InnerText!);
                var row = RowValue(spartLine.ReferenceSequence?.InnerText!);

                var newRef = $"{col}{row + shift}";
                spartLine.ReferenceSequence = new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(newRef);
            }
        }

        public static void MoveCellAnchor(WorksheetPart worksheetPart, int shift, string originArea)
        {
            // 所有位于 originArea 以后的描点
            var minRow = RowValue(originArea.Split(':')[0]);
            var maxRow = RowValue(originArea.Split(':')[1]);

            // OneCellAnchor
            var oneCells = worksheetPart?.DrawingsPart?.WorksheetDrawing.Descendants<OneCellAnchor>().ToList();
            if (oneCells != null)
            {
                foreach (var anchor in oneCells)
                {
                    if (anchor.FromMarker?.RowId?.InnerText == null) break;
                    var from = int.Parse(anchor.FromMarker?.RowId?.InnerText!);
                    if (from + 1 <= maxRow) break;
                    anchor.FromMarker!.RowId = new RowId((from + shift).ToString());
                }
            }

            var twoCells = worksheetPart?.DrawingsPart?.WorksheetDrawing.Descendants<TwoCellAnchor>().ToList();
            if(twoCells != null)
            {
                foreach (var anchor in twoCells)
                {
                    if (anchor.FromMarker?.RowId?.InnerText == null) break;
                    if (anchor.ToMarker?.RowId?.InnerText == null) break;

                    // 原始范围
                    var from = int.Parse(anchor.FromMarker?.RowId?.InnerText!);
                    var to = int.Parse(anchor.ToMarker?.RowId?.InnerText!);

                    // 在表格上方的不处理
                    // 横跨表格的也不处理（这种情况可能会导致显示出问题）

                    // 仅处理位于表格下方的
                    if (from + 1 >= maxRow)
                    {
                        anchor.FromMarker!.RowId = new RowId((from + shift).ToString());
                        anchor.ToMarker!.RowId = new RowId((to + shift).ToString());
                    }
                }
            }
        }

        /// <summary>
        /// 更新公式引用（仅处理行）
        /// </summary>
        /// <param name="formula">原先的公式</param>
        /// <param name="rowShift">偏移行数</param>
        private static void UpdateReference(Formula formula, int rowShift)
        {
            var parent = formula.Parent!;
            var fText = formula.InnerText;

            var newFormula = new Formula(ExpandRowReference(fText, rowShift));
            if (parent is StringReference)
            {
                var reference = parent as StringReference;
                if (reference != null)
                {
                    reference.Formula = newFormula;
                    reference.StringCache = null;
                }
            }
            else if (parent is NumberReference)
            {
                var reference = parent as NumberReference;
                if (reference != null)
                {
                    reference.Formula = newFormula;
                    reference.NumberingCache = null;
                }
            }
            else if (parent is MultiLevelStringReference)
            {
                Console.WriteLine("暂不支持 MultiLevelStringReference 类型引用");
            }
        }
    }
}
