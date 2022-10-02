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
            var outFile = "table-out.xlsx";
            TemplateGO.Render(R.FullPath("data/table.xlsx"), R.JsonFromFile("data/table.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }

        // 简单填充
        [TestMethod()]
        public void TableTestSimple()
        {
            var outFile = "table-simple-out.xlsx";
            TemplateGO.Render(R.FullPath("data/table-simple.xlsx"), R.JsonFromFile("data/table.json"), outFile);

            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);
        }

        // 纯数据数组
        [TestMethod()]
        public void TableTestRawData()
        {
            var outFile = "table-raw-out.xlsx";
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
