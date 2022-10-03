using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using TemplateGO.Parser;
using TemplateGO.Utils;

namespace TemplateGO.Processor
{
    internal class ProcTable : BaseProcess, IProcessor
    {
        void IProcessor.Process(ProcessParams p)
        {
            // 设置单元格内容（清空）
            SetCellValue(p.Cell, p.OriginValue, p.Parser, null, p.SharedStringTable);

            // 数据中指定的值
            var value = GetListValue(p);

            // 选项
            var options = new TableOptions(p.Parser.Options);

            // value 不能为空且
            var config = GetTableConfig(p, options);

            // 填充数据
            FillData(value, p, config, options);
        }

        private JsonElement? GetListValue(ProcessParams p)
        {
            var value = GetValueByProperty(p.Data, p.Parser.Property);
            if (value == null) return null;

            if (value is not JsonElement)
            {
                throw new ArgumentException($"{p.Parser.Property} 属性值不是数组对象，暂不支持处理。");
            }
            var json = value as JsonElement?;
            if (json?.ValueKind != JsonValueKind.Array)
            {
                throw new ArgumentException($"{p.Parser.Property} 属性值不是数组对象，暂不支持处理。");
            }

            return json;
        }

        private static TableConfig GetTableConfig(ProcessParams p, TableOptions options)
        {
            var tableConfig = new TableConfig();

            // 从当前行往下一行开始读取
            var curRowIndex = (p.Cell.Parent as Row)?.RowIndex?.Value;
            if (curRowIndex == null) throw new ArgumentException("模板配置错误");

            var row = p.WorksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex?.Value == curRowIndex + 1).FirstOrDefault();
            if (row == null) throw new ArgumentException($"模板配置错误，请在 {p.Parser.Origin} 下一行设置属性列");

            var beginColumn = CellUtils.ColumnValue(p.Cell.CellReference!);
            var cells = row.Descendants<Cell>().Where(cell =>
                CellUtils.ColumnValue(cell.CellReference!) >= beginColumn
            ).ToList();

            foreach (var cell in cells)
            {
                var value = CellUtils.GetCellString(cell, p.SharedStringTable);
                tableConfig.KeyMap.Add(CellUtils.ColumnReference(cell.CellReference!), value);
            }

            // 数据开始行编号 当前Cell + 列属性 + 标题行数目
            tableConfig.BeginRowIndex = curRowIndex.Value + 2 + options.TitleCount;

            // 当前数据行数
            tableConfig.ExistRowCount = (uint)p.WorksheetPart.Worksheet.Descendants<Row>().Count((r) => r.RowIndex?.Value >= tableConfig.BeginRowIndex);

            // 检查数据区域是否有表格
            foreach (var tablePart in p.WorksheetPart.TableDefinitionParts)
            {
                var table = tablePart.Table;
                var tabRef = table.Reference;
                if (table == null || tabRef == null) continue;

                var range = tabRef?.Value?.Split(':');
                if (range?.Length != 2)
                {
                    Console.WriteLine($"无法解析表格 {table.Name} 引用区域。");
                    continue;
                }

                // 开始行 = BeginRowIndex - 1（表格标题）
                if (CellUtils.RowValue(range[0]) != tableConfig.BeginRowIndex - 1)
                {
                    continue;
                }

                // 列范围只限定有重合即可
                if (CellUtils.IsColumnIntersect(tabRef!, $"{cells[0].CellReference}:{cells.Last().CellReference}"))
                {
                    if (tableConfig.RefTable != null)
                    {
                        throw new ArgumentException("暂不支持多个表格");
                    }
                    tableConfig.RefTable = table;
                }
            }

            // 有表格时样本数量通过表格计算
            tableConfig.SampleCount = options.SampleCount;
            if (tableConfig.RefTable != null)
            {
                var tabRef = tableConfig.RefTable.Reference?.Value;
                if (tabRef == null) throw new ArgumentException($"{tableConfig.RefTable.Name} 中引用无效。");
                var tabRefRanges = tabRef.Split(':');
                var rows = CellUtils.RowValue(tabRefRanges[1]) - CellUtils.RowValue(tabRefRanges[0]);
                if (tableConfig.RefTable.TotalsRowCount is not null)
                {
                    rows -= Convert.ToInt32(tableConfig.RefTable.TotalsRowCount.Value);
                }

                tableConfig.SampleCount = (uint)rows;
            }

            // 总数据行数
            return tableConfig;
        }

        private struct TableConfig
        {
            public TableConfig() { }

            /// <summary>
            /// key为列名，如A/B/C等
            /// value为属性名
            /// </summary>
            public Dictionary<string, string> KeyMap = new();

            /// <summary>
            /// 开始行，下标从1开始 如 A1 表示第1行
            /// </summary>
            public uint BeginRowIndex { get; set; } = 0;

            /// <summary>
            /// 现有数据行数
            /// </summary>
            public uint ExistRowCount { get; set; } = 0;

            /// <summary>
            /// 引用的表格
            /// </summary>
            public Table? RefTable { get; set; } = null;

            /// <summary>
            /// 样例数据条数
            /// </summary>
            public uint SampleCount { get; set; } = 0;
        }


        private void FillData(JsonElement? data, ProcessParams p, TableConfig config, TableOptions options)
        {
            // 空数据视为 []
            data ??= JsonDocument.Parse(@"[]")!.RootElement;

            uint idx = 0, rowIndex = 0;
            var sheetData = p.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            // 使用第1条样例数据的格式
            var refRow = p.WorksheetPart.Worksheet.Descendants<Row>().FirstOrDefault((r) => r.RowIndex?.Value == config.BeginRowIndex);
            var list = data.Value.EnumerateArray().ToList();

            // 只处理选项中指定 Limit 条数
            if (options.Limit != null)
            {
                list = list.Take(options.Limit.Value).ToList();
            }

            // 插入的游标位置(从上往下找到离 BeginRowIndex 最近的一条)
            var latestRow = p.WorksheetPart.Worksheet.Descendants<Row>().LastOrDefault((r) => r.RowIndex?.Value < config.BeginRowIndex);
            if (latestRow == null) throw new Exception("模板格式错误，无法定位数据插入位置。");
            foreach (var value in list)
            {
                // 行
                var row = (Row?)refRow?.Clone() ?? new Row();
                rowIndex = config.BeginRowIndex + idx;
                row.RowIndex = rowIndex;
                sheetData.InsertAfter(row, latestRow);
                latestRow = row;

                foreach (var col in config.KeyMap)
                {
                    var columnRef = col.Key;
                    var propName = col.Value;

                    // 内容
                    object? cellValue = GetCellValue(value, propName, idx, rowIndex);
                    var isFormula = propName.StartsWith("=");

                    // 设置单元格内容
                    var cellReference = $"{columnRef}{rowIndex}";
                    var cell = row.Descendants<Cell>().Where(c => CellUtils.ColumnReference(c.CellReference!) == columnRef).FirstOrDefault();
                    if (cell == null)
                    {
                        cell = new Cell();
                        row.Append(cell);
                    }
                    cell.CellReference = cellReference;

                    if (isFormula)
                    {
                        cell.CellFormula = new CellFormula($"{cellValue}");
                        cell.CellValue = null;
                    }
                    else
                    {
                        SetCellValueWithType(cell, cellValue);
                    }
                }

                // 如果该行是复制来的，则清除数据区域以外的内容并更新引用
                if (refRow != null)
                {
                    var cells = row.Descendants<Cell>().ToList();
                    if (cells.FirstOrDefault() != null &&
                        CellUtils.RowValue(cells.FirstOrDefault()!.CellReference!) != rowIndex)
                    {
                        foreach (var cell in row.Descendants<Cell>())
                        {
                            cell.CellReference = $"{CellUtils.ColumnReference(cell.CellReference!)}{rowIndex}";

                            // 清除内容
                            cell.CellValue = null;
                            cell.CellFormula = null;
                            cell.DataType = null;
                        }
                    }
                }
                idx++;
            }

            // 最后一条数据行号
            var latestIndex = latestRow.RowIndex!.Value;

            // latestIndex 及其后的数据均需要移动
            var allRows = p.WorksheetPart.Worksheet.Descendants<Row>().ToList();
            var latestIdx = allRows.IndexOf(latestRow);
            if (latestIdx == -1) throw new Exception("插入表格后无法处理后续数据");
            var afterLatestList = allRows.Skip(latestIdx + 1).ToList();

            // 更新后续表格引用
            var shift = Convert.ToInt32(idx) - Convert.ToInt32(config.SampleCount);

            // 删除样本数据
            if (config.SampleCount > 0)
            {
                // 样本数据（取数组中 latestRow 后面的 SampleCount 条）
                var sampleRows = afterLatestList.Take((int)config.SampleCount).ToList();
                afterLatestList = afterLatestList.Skip(sampleRows.Count).ToList();

                // 表格需要保留至少1行数据
                if (config.RefTable != null && idx <= 0)
                {
                    var first = sampleRows[0];
                    foreach (var cell in first.Descendants<Cell>())
                    {
                        cell.CellFormula = null;
                        cell.CellValue = null;
                        cell.DataType = null;
                    }
                    sampleRows.RemoveAt(0);
                    shift++;
                }
                foreach (var row in sampleRows) row.Remove();
            }

            // 更新latestRow后面几行数据引用
            if (afterLatestList.Count > 0)
            {
                foreach (var row in afterLatestList)
                {
                    var oldIdx = row.RowIndex!.Value;
                    var newIdx = Convert.ToUInt32(oldIdx + shift);
                    row.RowIndex = newIdx;
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        cell.CellReference = $"{CellUtils.ColumnReference(cell.CellReference!)}{newIdx}";
                    }
                }
            }

            // 如果是表格则更新表格引用
            if (config.RefTable != null)
            {
                var table = config.RefTable;
                if (!string.IsNullOrEmpty(config.RefTable.Reference))
                {
                    table.Reference = CellUtils.UpdateToReference(table.Reference!, shift);

                    // AutoFilter 可能不存在
                    if (table.AutoFilter != null)
                    {
                        table.AutoFilter.Reference = CellUtils.UpdateToReference(table.AutoFilter.Reference!, shift);
                    }
                }

                // 如果有汇总行则删除公式计算的内容
                if (config.RefTable.TotalsRowCount?.Value > 0 && latestRow != null)
                {
                    var totalRow = p.WorksheetPart.Worksheet.Descendants<Row>().FirstOrDefault((r) => r.RowIndex?.Value == latestRow.RowIndex!.Value + 1);
                    if (totalRow != null)
                    {
                        foreach (var cell in totalRow.Descendants<Cell>())
                        {
                            if (cell.CellFormula != null) cell.CellValue = null;
                        }
                    }
                }
            }

            // 之前的表格区域
            var range = $"{config.KeyMap.Keys.First()}{config.BeginRowIndex}:{config.KeyMap.Keys.Last()}{config.BeginRowIndex+config.SampleCount}";

            // 更新图表引用
            CellUtils.UpdateChartReference(p.WorkbookPart, p.Sheet.Name!, shift, range);

            // 更新迷你图
            CellUtils.UpdateSparklineArea(p.WorksheetPart, shift, range);
            CellUtils.UpdateSparklineReference(p.WorkbookPart, p.Sheet.Name!, shift, range);
        }

        private static object? GetCellValue(JsonElement data, string propName, uint index, uint rowNumber)
        {
            // 内容
            object? cellValue = null;
            switch (propName)
            {
                case "-":
                    break;
                case "#index":
                    cellValue = index;
                    break;
                case "#seq":
                    cellValue = index + 1;
                    break;
                case "#value":
                    cellValue = JsonUtils.GetValue(data);
                    break;

                default:
                    {
                        if (propName.StartsWith("=:"))
                        {
                            var formual = propName[2..];
                            formual = formual.Replace("#row", $"{rowNumber}");
                            formual = formual.Replace("#index", $"{index}");
                            formual = formual.Replace("#seq", $"{index + 1}");
                            return formual.Trim();
                        }
                        else if (propName.StartsWith("="))
                        {
                            propName = propName[1..].Trim();
                        }

                        try
                        {
                            cellValue = JsonUtils.GetValue(data, propName);
                        }
                        catch { }
                    }
                    break;
            }
            return cellValue;
        }
    }
}
