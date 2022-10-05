using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using TemplateGO.Utils;
using TemplateGOTests;

namespace TemplateGO.Tests
{
    [TestClass()]
    public class TemplateRenderTests
    {
        [TestMethod()]
        public void RenderTestXlsx()
        {
            var outFile = R.OutFullPath("go-out-xlsx.xlsx");
            TemplateRender.Render(R.FullPath("data/go.xlsx"), R.JsonFromFile("data/go.json"), outFile);
            AssertFile(outFile);
        }

        [TestMethod()]
        public void RenderTestXltx()
        {
            var outFile = R.OutFullPath("go-out-xltx.xlsx");
            TemplateRender.Render(R.FullPath("data/go.xltx"), R.JsonFromFile("data/go.json"), outFile);
            AssertFile(outFile);
        }

        private static void AssertFile(string outFile)
        {
            // 应该能打开文档
            using var doc = SpreadsheetDocument.Open(outFile, false);
            Assert.IsNotNull(doc);
            Assert.IsNotNull(doc.WorkbookPart);
            var sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>();
            Assert.IsNotNull(sheets);

            var ssTable = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            var siPart = doc.WorkbookPart?.WorksheetParts.GetEnumerator();
            var cellValues = new Dictionary<string, string>() {
                    {"B3", "来自测试数据" },
                    {"B4", "来自测试数据" },
                    {"B5", "${hello|nono}" }
                };
            while (siPart != null && siPart.MoveNext())
            {
                var cells = siPart.Current.Worksheet.Descendants<Cell>();
                foreach (var cell in cells)
                {
                    if (!cellValues.ContainsKey(cell.CellReference!)) continue;

                    var value = CellUtils.GetCellString(cell, ssTable);
                    Assert.AreEqual(cellValues[cell.CellReference!], value);
                }
            }
        }
    }
}