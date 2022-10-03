using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using TemplateGO.Utils;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class ProcTableTests
    {
        // 表格
        [TestMethod()]
        public void TableTest()
        {
            var outFile = R.OutFullPath("table-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table.xlsx"), R.JsonFromFile("data/table.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }

        // 图表测试
        [TestMethod()]
        public void TableTestChart()
        {
            var outFile = R.OutFullPath("table-chart-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table-chart.xlsx"), R.JsonFromFile("data/table.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }

        // 空数据
        [TestMethod()]
        public void TableTestEmpty()
        {
            var outFile = R.OutFullPath("table-empty-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table.xlsx"), R.JsonFromFile("data/table-empty.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 表格测试中数据应该存在
            Assert.AreEqual("汇总", R.CellStringValue(doc, "表格测试", "A8"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "表格测试", "A10"));

            Assert.AreEqual("下方内容不受影响", R.CellStringValue(doc, "无汇总行", "D12"));
        }
        // 空数据
        [TestMethod()]
        public void TableTestSimpleEmpty()
        {
            var outFile = R.OutFullPath("table-simple-empty-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table-simple.xlsx"), R.JsonFromFile("data/table-empty.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 检查单元格内容
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "有样本数据", "A6"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "有样本数据但sampleCount为0", "A11"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "无样本数据", "A7"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "仅有样式样本数据", "A6"));
        }

        // 大数据
        [TestMethod()]
        public void TableTestLarge()
        {
            var json = R.JsonFromFile("data/table-large.json");
            var outFile = R.OutFullPath("table-large-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table-large.xlsx"), json, outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 所有人名
            var names = new List<string>();
            foreach(var item in json.GetProperty("data").EnumerateArray())
            {
                names.Add(item.GetProperty("name").ToString());
            }

            // 使用表格!D4:D203所有数据
            var names1 = new List<string>();
            for(var row=4; row<=203; row++)
            {
                names1.Add(R.CellStringValue(doc, "使用表格", $"D{row}")!);
            }
            Assert.IsTrue(names.SequenceEqual(names1));

            // 汇总内容也存在
            Assert.AreEqual("汇总", R.CellStringValue(doc, "使用表格", "A204"));
            Assert.AreEqual("SUBTOTAL(103,表1[姓名])", R.CellStringValue(doc, "使用表格", "D204"));

            // 不使用表格!D4:D203所有数据
            var names2 = new List<string>();
            for (var row = 4; row <= 203; row++)
            {
                names2.Add(R.CellStringValue(doc, "不使用表格", $"D{row}")!);
            }
            Assert.IsTrue(names.SequenceEqual(names2));
        }

        // 简单填充
        [TestMethod()]
        public void TableTestSimple()
        {
            var outFile = R.OutFullPath("table-simple-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table-simple.xlsx"), R.JsonFromFile("data/table.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 检查单元格内容
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "有样本数据", "A16"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "有样本数据但sampleCount为0", "A21"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "无样本数据", "A17"));
            Assert.AreEqual("这是其它的一些内容", R.CellStringValue(doc, "仅有样式样本数据", "A16"));
        }

        // 纯数据数组
        [TestMethod()]
        public void TableTestRawData()
        {
            var outFile = R.OutFullPath("table-raw-out.xlsx");
            TemplateGO.Render(R.FullPath("data/table-raw.xlsx"), R.JsonFromFile("data/table-raw.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            // 检查D列数据是否正常
            var sheetPart = doc.WorkbookPart.GetPartById(sheets.FirstOrDefault()!.Id!) as WorksheetPart;
            var shareStringTable = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable;
            Assert.IsNotNull(sheetPart);

            var users = new List<string>()
            {
                "贾阳煦", "钟惠玲", "孟三春", "边念蕾", "终苏凌", "勾融雪", "甄映寒", "濮高寒", "范弘致", "公冰洁"
            };
            var cells = sheetPart.Worksheet.Descendants<Cell>().Where(r => r.CellReference!.Value!.StartsWith("D") && (r.Parent as Row)!.RowIndex!.Value >= 5);
            Assert.AreEqual(users.Count, cells.Count());
            foreach (var cell in cells)
            {
                var cellval = CellUtils.GetCellString(cell, shareStringTable);
                Assert.IsTrue(users.IndexOf(cellval) >= 0);
            }
        }
    }
}
